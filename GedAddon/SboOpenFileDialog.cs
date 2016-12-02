using System;
using System.Drawing;
using SAPbouiCOM;
using MSComctlLib;
using DocMageFramework.AppUtils;
using DocMageFramework.FileUtils;


namespace GedAddon
{
    internal sealed class PictureHelper : System.Windows.Forms.AxHost
    {
        private PictureHelper() : base(string.Empty) { }
        /// <summary>
        /// Convert the image to an ipicturedisp.
        /// </summary>
        public new static Object GetIPictureDispFromPicture(Image image)
        {
            return System.Windows.Forms.AxHost.GetIPictureDispFromPicture(image);
        }
        /// <summary>
        /// Convert the dispatch interface into an image object.
        /// </summary>
        public new static Image GetPictureFromIPicture(Object picture)
        {
            return System.Windows.Forms.AxHost.GetPictureFromIPicture(picture);
        }
    }

    public class SboOpenFileDialog
    {
        private SAPbouiCOM.Application sboApplication;

        private String parentForm;

        private SAPbouiCOM.Form openDialog;

        private String dialogUID;

        private String filter;

        private String initialDirectory;

        private TreeView treeView;

        private String lastError;

        private String fileName;

        public String FileName
        {
            get { return fileName; }
        }


        public SboOpenFileDialog(SAPbouiCOM.Application sboApplication, String parentForm, String filter)
        {
            this.sboApplication = sboApplication;
            this.parentForm = parentForm;
            this.filter = filter;
            this.initialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        }

        public void ShowDialog()
        {
            if (sboApplication == null) return;

            // Gera um identificador único para o dialogo
            String dialogUID = Cipher.GenerateHash(DateTime.Now.ToString("HH:mm:ss.fff"));

            FormCreationParams formParams = (FormCreationParams)(sboApplication.CreateObject(BoCreatableObjectType.cot_FormCreationParams));
            formParams.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            formParams.FormType = "openDialog";
            formParams.UniqueID = dialogUID;

            // Cria o form de escolha de arquivo
            try
            {
                openDialog = sboApplication.Forms.AddEx(formParams);
                openDialog.AutoManaged = false;
                openDialog.Title = "Seleção de Arquivo";
                openDialog.Top = 150;
                openDialog.Left = 450;
                openDialog.ClientWidth = 480;
                openDialog.ClientHeight = 320;
                AddControls();
                AddData();
                openDialog.Visible = true;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void AddControls()
        {
            SAPbouiCOM.Item label1 = openDialog.Items.Add("lblPath", BoFormItemTypes.it_STATIC);
            label1.Left = 10;
            label1.Top = 10;
            label1.Width = 60;
            label1.Height = 25;
            ((SAPbouiCOM.StaticText)(label1.Specific)).Caption = "Diretório: ";

            SAPbouiCOM.Item pathItem = openDialog.Items.Add("txtPath", BoFormItemTypes.it_EDIT);
            pathItem.Left = 60;
            pathItem.Top = 10;
            pathItem.Width = 300;
            pathItem.Height = 25;
            pathItem.Enabled = false;
            SAPbouiCOM.EditText pathSpecific = (SAPbouiCOM.EditText)pathItem.Specific;
            pathSpecific.Value = initialDirectory;

            SAPbouiCOM.Item treeItem = openDialog.Items.Add("imgTree", BoFormItemTypes.it_ACTIVE_X);
            treeItem.Left = 10;
            treeItem.Top = 40;
            treeItem.Width = 460;
            treeItem.Height = 200;
            ((SAPbouiCOM.ActiveX)treeItem.Specific).ClassID = "MSComctlLib.TreeCtrl.2";
            treeView = (TreeView)((SAPbouiCOM.ActiveX)treeItem.Specific).Object;
            treeView.Scroll = true;
            Image picture = Bitmap.FromFile(FileResource.MapDesktopResource(@"Images\Folder.png"));
            Object pictureIndex = 0;
            Object pictureKey = "Folder";
            Object pictureRef = PictureHelper.GetIPictureDispFromPicture(picture);
            ImageList imageList = new ImageListClass();
            imageList.ListImages.Add(ref pictureIndex, ref pictureKey, ref pictureRef);
            treeView.ImageList = (Object)imageList;
            Object noValue = Type.Missing;
            Object imageIndex = 0;
            Object rootKey = @"C:\Program Files";
            Object rootName = @"C:\Program Files";
            Object treeNode = treeView.Nodes.Add(ref noValue, ref noValue, ref rootKey, ref rootName, ref imageIndex, ref noValue);
            Object leafKey = @"C:\Work";
            Object leafName = @"C:\Work";
            treeView.Nodes.Add(ref treeNode, ref noValue, ref leafKey, ref leafName, ref imageIndex, ref noValue);

            SAPbouiCOM.Item label2 = openDialog.Items.Add("lblFile", BoFormItemTypes.it_STATIC);
            label2.Left = 10;
            label2.Top = 250;
            label2.Width = 60;
            label2.Height = 25;
            ((SAPbouiCOM.StaticText)(label2.Specific)).Caption = "Arquivo: ";

            SAPbouiCOM.Item fileItem = openDialog.Items.Add("txtFile", BoFormItemTypes.it_EDIT);
            fileItem.Left = 60;
            fileItem.Top = 250;
            fileItem.Width = 300;
            fileItem.Height = 25;
            fileItem.Enabled = false;
            SAPbouiCOM.EditText fileSpecific = (SAPbouiCOM.EditText)fileItem.Specific;
            fileSpecific.Value = "Apostila.pdf";

            SAPbouiCOM.Item label3 = openDialog.Items.Add("lblFilter", BoFormItemTypes.it_STATIC);
            label3.Left = 10;
            label3.Top = 280;
            label3.Width = 60;
            label3.Height = 25;
            ((SAPbouiCOM.StaticText)(label3.Specific)).Caption = "Tipo: ";

            SAPbouiCOM.Item filterItem = openDialog.Items.Add("txtFilter", BoFormItemTypes.it_EDIT);
            filterItem.Left = 60;
            filterItem.Top = 280;
            filterItem.Width = 300;
            filterItem.Height = 25;
            filterItem.Enabled = false;
            SAPbouiCOM.EditText filterSpecific = (SAPbouiCOM.EditText)filterItem.Specific;
            filterSpecific.Value = this.filter;

            SAPbouiCOM.Item okButton = openDialog.Items.Add("btnOk", BoFormItemTypes.it_BUTTON);
            okButton.Left = 380;
            okButton.Top = 250;
            okButton.Width = 80;
            okButton.Height = 25;
            ((SAPbouiCOM.Button)(okButton.Specific)).Caption = "OK";

            SAPbouiCOM.Item cancelButton = openDialog.Items.Add("btnCancel", BoFormItemTypes.it_BUTTON);
            cancelButton.Left = 380;
            cancelButton.Top = 280;
            cancelButton.Width = 80;
            cancelButton.Height = 25;
            ((SAPbouiCOM.Button)(cancelButton.Specific)).Caption = "Cancel";
        }

        private void AddData()
        {
            UserDataSource parentFormDs = openDialog.DataSources.UserDataSources.Add("parentForm", BoDataType.dt_SHORT_TEXT, 50);
            parentFormDs.Value = this.parentForm;
        }
    }

}
