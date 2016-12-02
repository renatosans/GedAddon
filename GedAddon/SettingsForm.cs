using System;
using System.IO;
using SAPbouiCOM;
using SystemIntegration;
using DocMageFramework.FileUtils;


namespace GedAddon
{
    public class SettingsForm
    {
        private SAPbouiCOM.Application sboApplication;

        private SAPbouiCOM.Form settingsForm;


        public SettingsForm(SAPbouiCOM.Application sboApplication)
        {
            this.sboApplication = sboApplication;
            AttachToForm();
            // não é necessário criar o objecto no SAP caso o form já esteja aberto
            if (settingsForm != null) return;

            TextReader textReader = new StreamReader(FileResource.MapDesktopResource(@"Xml\GedSettings.srf"));
            String formXml = textReader.ReadToEnd();
            FormCreationParams creationPackage = (FormCreationParams)sboApplication.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "frmSettings";
            creationPackage.FormType = "customForm";
            creationPackage.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            creationPackage.XmlData = formXml;
            settingsForm = sboApplication.Forms.AddEx(creationPackage);
        }

        private void AttachToForm()
        {
            try { settingsForm = sboApplication.Forms.Item("frmSettings"); }
            catch { settingsForm = null; }
        }

        public void LoadSettings()
        {
            GedSettings settings = new GedSettings();
            settings.LoadFromXml();
            if (settings.LastError != null)
            {
                // Sai do método caso não consiga ler o arquivo de configurações
                sboApplication.MessageBox(settings.LastError, 1, "Ok", "", "");
                return;
            }

            if (settingsForm == null) return;

            SAPbouiCOM.Item serverItem = settingsForm.Items.Item("txtServer");
            SAPbouiCOM.EditText txtServer = ((SAPbouiCOM.EditText)(serverItem.Specific));
            SAPbouiCOM.Item docLibItem = settingsForm.Items.Item("txtDocLib");
            SAPbouiCOM.EditText txtDocLib = ((SAPbouiCOM.EditText)(docLibItem.Specific));
            SAPbouiCOM.Item userItem = settingsForm.Items.Item("txtUser");
            SAPbouiCOM.EditText txtUser = ((SAPbouiCOM.EditText)(userItem.Specific));
            SAPbouiCOM.Item passItem = settingsForm.Items.Item("txtPass");
            SAPbouiCOM.EditText txtPass = ((SAPbouiCOM.EditText)(passItem.Specific));

            txtServer.Value = settings.Server;
            txtDocLib.Value = settings.Library;
            txtUser.Value = settings.Username;
            txtPass.Value = settings.Password;
        }

        public void SaveSettings()
        {
            if (settingsForm == null) return;

            SAPbouiCOM.Item serverItem = settingsForm.Items.Item("txtServer");
            SAPbouiCOM.EditText txtServer = ((SAPbouiCOM.EditText)(serverItem.Specific));
            SAPbouiCOM.Item docLibItem = settingsForm.Items.Item("txtDocLib");
            SAPbouiCOM.EditText txtDocLib = ((SAPbouiCOM.EditText)(docLibItem.Specific));
            SAPbouiCOM.Item userItem = settingsForm.Items.Item("txtUser");
            SAPbouiCOM.EditText txtUser = ((SAPbouiCOM.EditText)(userItem.Specific));
            SAPbouiCOM.Item passItem = settingsForm.Items.Item("txtPass");
            SAPbouiCOM.EditText txtPass = ((SAPbouiCOM.EditText)(passItem.Specific));

            GedSettings settings = new GedSettings();
            settings.SaveToXml(txtServer.Value, txtDocLib.Value, txtUser.Value, txtPass.Value);
            if (settings.LastError != null)
                sboApplication.MessageBox(settings.LastError, 1, "Ok", "", "");
            settingsForm.Close();
        }
    }

}
