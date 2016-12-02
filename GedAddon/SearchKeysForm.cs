using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using SAPbouiCOM;
using SystemIntegration;
using DocMageFramework.FileUtils;


namespace GedAddon
{
    public class SearchKeysForm
    {
        private SAPbouiCOM.Application sboApplication;

        private SAPbouiCOM.Form searchKeysForm;


        public SearchKeysForm(SAPbouiCOM.Application sboApplication)
        {
            this.sboApplication = sboApplication;
            AttachToForm();
            // não é necessário criar o objecto no SAP caso o form já esteja aberto
            if (searchKeysForm != null) return;

            TextReader textReader = new StreamReader(FileResource.MapDesktopResource(@"Xml\GedSearchKeys.srf"));
            String formXml = textReader.ReadToEnd();
            formXml = formXml.Replace("Logo.png", FileResource.MapDesktopResource(@"Images\Logo.png"));

            FormCreationParams creationPackage = (FormCreationParams)sboApplication.CreateObject(BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "frmSearchKeys";
            creationPackage.FormType = "customForm";
            creationPackage.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            creationPackage.XmlData = formXml;
            searchKeysForm = sboApplication.Forms.AddEx(creationPackage);
        }

        private void AttachToForm()
        {
            try { searchKeysForm = sboApplication.Forms.Item("frmSearchKeys"); }
            catch { searchKeysForm = null; }
        }

        /// <summary>
        /// Carrega o combo com os nomes dos campos da biblioteca de documentos
        /// </summary>
        public void LoadComboValues(NameValueCollection fieldNames)
        {
            if (searchKeysForm == null) return;

            SAPbouiCOM.Item cmbFieldsItem = searchKeysForm.Items.Item("cmbFields");
            SAPbouiCOM.ComboBox cmbFields = (SAPbouiCOM.ComboBox)cmbFieldsItem.Specific;
            // Remove os items do combo
            while (cmbFields.ValidValues.Count > 0)
                cmbFields.ValidValues.Remove(cmbFields.ValidValues.Count - 1, BoSearchKey.psk_Index);
            // Preenche novamente o combo
            foreach (String fieldName in fieldNames)
                cmbFields.ValidValues.Add(fieldName, fieldNames[fieldName]);
        }

        /// <summary>
        /// Salva a lista de chaves no XML
        /// </summary>
        public void SaveSearchKeys()
        {
            if (searchKeysForm == null) return;

            SAPbouiCOM.Item lstChoosenItem = searchKeysForm.Items.Item("lstChoosen");
            SAPbouiCOM.EditText lstChoosen = (SAPbouiCOM.EditText)lstChoosenItem.Specific;

            String[] controlValues = lstChoosen.Value.Split(new String[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            List<String> choosenKeys = new List<String>();
            choosenKeys.AddRange(controlValues);
            GedSearchKeys gedSearchKeys = new GedSearchKeys();
            gedSearchKeys.SaveToXml(choosenKeys);
            if (gedSearchKeys.LastError != null)
                sboApplication.MessageBox(gedSearchKeys.LastError, 1, "Ok", "", "");
            searchKeysForm.Close();
        }

        /// <summary>
        /// Adiciona a chave selecionada no combo a lista de chaves
        /// </summary>
        public void AddChoosenKey()
        {
            if (searchKeysForm == null) return;

            SAPbouiCOM.Item lstChoosenItem = searchKeysForm.Items.Item("lstChoosen");
            SAPbouiCOM.EditText lstChoosen = (SAPbouiCOM.EditText)lstChoosenItem.Specific;

            String[] controlValues = lstChoosen.Value.Split(new String[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            List<String> choosenKeys = new List<String>();
            choosenKeys.AddRange(controlValues);
            if (choosenKeys.Count >= 2)
            {
                sboApplication.MessageBox("O limite de campos de busca foi excedido", 1, "Ok", "", "");
                return;
            }

            SAPbouiCOM.Item cmbFieldsItem = searchKeysForm.Items.Item("cmbFields");
            SAPbouiCOM.ComboBox cmbFields = (SAPbouiCOM.ComboBox)cmbFieldsItem.Specific;
            if (!choosenKeys.Contains(cmbFields.Selected.Value))
                choosenKeys.Add(cmbFields.Selected.Value); // adiciona o item selecionado(combobox)

            String concatChoosenKeys = null;
            foreach (String key in choosenKeys)
            {
                if (concatChoosenKeys != null) concatChoosenKeys += Environment.NewLine;
                concatChoosenKeys += key;
            }
            lstChoosen.Value = concatChoosenKeys;
        }
    }

}
