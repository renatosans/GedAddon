using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using SAPbouiCOM;
using SystemIntegration;
using DocMageFramework.AppUtils;
using DocMageFramework.FileUtils;
using Microsoft.SharePoint.Client;


namespace GedAddon
{
    /// <summary>
    /// Classe responsável pelo controle da interface do usuário renderizando os controles
    /// e recebendo eventos, a regra de negócio é delegada/encapsulada em GedContext
    /// </summary>
    public class GuiController: IGraphicalDisplay
    {
        private const int GED_TAB = 11; // Identificador da aba GED

        private SAPbouiCOM.Application sboApplication;

        private SearchPanel searchPanel;

        private SboFlatButton uploadButton;

        private SboFlatButton searchButton;

        private GedContext gedContext;

        private String lastError;


        /// <summary>
        /// Construtor da classe, primeiramente faz a conexão com o servidor para depois acoplar a
        /// interface do SAP B1, isso é necessário pois o SAP precisa de um tempo para disponibilizar
        /// a interface e também porque é necessário buscar no servidor os campos de busca antes de
        /// de montar a interface
        /// </summary>
        public GuiController()
        {
            this.sboApplication = null; // permanece nulo até InitializeGui ser chamado
            this.gedContext = new GedContext(this);
            this.gedContext.Connect();
        }

        public void InitializeGui(SAPbouiCOM.Application sboApplication)
        {
            if (sboApplication == null) return;

            this.sboApplication = sboApplication;
            AddMenuItems();
        }

        private void AddMenuItems()
        {
            SAPbouiCOM.Menus sboMenus = sboApplication.Menus;
            SAPbouiCOM.MenuItem sboMenuItem = sboApplication.Menus.Item("43520");

            MenuCreationParams creationParams = (MenuCreationParams)(sboApplication.CreateObject(BoCreatableObjectType.cot_MenuCreationParams));
            creationParams.Type = BoMenuType.mt_POPUP;
            creationParams.UniqueID = "GedMenu";
            creationParams.String = "Gerenciamento de Documentos";
            creationParams.Image = FileResource.MapDesktopResource(@"Images\GED.png");
            creationParams.Position = sboMenuItem.SubMenus.Count + 1;

            sboMenus = sboMenuItem.SubMenus;
            try
            {
                sboMenus.AddEx(creationParams);
                sboMenuItem = sboApplication.Menus.Item("GedMenu");
                sboMenus = sboMenuItem.SubMenus;

                creationParams.Type = BoMenuType.mt_STRING;
                creationParams.UniqueID = "GedMenuItem1";
                creationParams.String = "Configurar Servidor"; // settings form
                sboMenus.AddEx(creationParams);

                creationParams.Type = BoMenuType.mt_STRING;
                creationParams.UniqueID = "GedMenuItem2";
                creationParams.String = "Configurar Busca"; // search keys form
                sboMenus.AddEx(creationParams);

                creationParams.Type = BoMenuType.mt_STRING;
                creationParams.UniqueID = "GedMenuItem3";
                creationParams.String = "Open File";
                sboMenus.AddEx(creationParams);
            }
            catch (Exception exc)
            {
                sboApplication.MessageBox(exc.Message, 1, "Ok", "", "");
            }
        }

        public void MenuEvent(ref MenuEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = false;
            switch (pVal.MenuUID)
            {
                case "GedMenuItem1":
                    SettingsForm settingsForm = new SettingsForm(sboApplication);
                    settingsForm.LoadSettings();
                    break;
                case "GedMenuItem2":
                    SearchKeysForm searchKeysForm = new SearchKeysForm(sboApplication);
                    searchKeysForm.LoadComboValues(gedContext.FieldNames);
                    break;
                case "GedMenuItem3":
                    SboOpenFileDialog openFileDialog = new SboOpenFileDialog(sboApplication, "gedForm", "*.PDF");
                    openFileDialog.ShowDialog();
                    break;
                default:
                    bubbleEvent = true; // permite que os outros menus do SAP respondam aos eventos
                    break;
            }
        }

        public void FormDataEvent(ref BusinessObjectInfo businessObjectInfo, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if ((businessObjectInfo.FormTypeEx == "134") && (businessObjectInfo.BeforeAction == false))
            {
                if (businessObjectInfo.EventType == BoEventTypes.et_FORM_DATA_LOAD)
                {
                    // Habilita os botões do GED quando um cliente é selecionado
                    if (uploadButton != null) uploadButton.SetState(true);
                    if (searchButton != null) searchButton.SetState(true);
                    // Limpa o grid de documentos para este cliente
                    ClearGrid(businessObjectInfo.FormUID, "gridDocs");
                }
            }
        }

        public void ItemEvent(String formUID, ref ItemEvent pVal, out bool bubbleEvent)
        {
            bubbleEvent = true;
            if ((pVal.FormType == 134) && (pVal.BeforeAction == false))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                    CreateGedTab(formUID);
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE)
                    ResizeGedTab(formUID);
                if ((pVal.ItemUID == "GedFolder") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    OpenGedTab(formUID);
                if (pVal.ItemUID == "5")
                    CardCodeChanged(formUID, pVal.EventType);
                if ((pVal.ItemUID == "imgSearch") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SearchButtonClicked(formUID);
                if ((pVal.ItemUID == "imgUpload") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    UploadButtonClicked(formUID);
                if ((pVal.ItemUID == "gridDocs") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK))
                    SelectCell(pVal.ColUID, pVal.Row, formUID);
            }
            if ((formUID == "frmSettings") && (pVal.ItemUID == "btnSave") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {
                bubbleEvent = false;
                // Salva as configurações no XML
                SettingsForm settingsForm = new SettingsForm(sboApplication);
                settingsForm.SaveSettings();
                // Reconecta ao servidor com as  novas configurações
                gedContext.Connect();
            }
            if ((formUID == "frmSearchKeys") && (pVal.ItemUID == "btnAdd") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {
                bubbleEvent = false;
                SearchKeysForm searchKeysForm = new SearchKeysForm(sboApplication);
                searchKeysForm.AddChoosenKey();
            }
            if ((formUID == "frmSearchKeys") && (pVal.ItemUID == "btnSave") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
            {
                bubbleEvent = false;
                // Salva as configurações no XML
                SearchKeysForm searchKeysForm = new SearchKeysForm(sboApplication);
                searchKeysForm.SaveSearchKeys();
                // Reconecta ao servidor após gravar as configurações
                gedContext.Connect();
            }
            if ((pVal.FormTypeEx == "openDialog") && (pVal.BeforeAction == false))
            {
                if ((pVal.ItemUID == "btnOk") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
                    FileSelected(formUID);
                if ((pVal.ItemUID == "btnCancel") && (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED))
                    FileNotSelected(formUID);
            }
        }

        private void CreateGedTab(String formUID)
        {
            try
            {
                // Insere a aba ao lado da última visivel
                SAPbouiCOM.Form bpForm = GetSBOForm(formUID);
                SAPbouiCOM.Item lastFolder = bpForm.Items.Item("9");
                SAPbouiCOM.Item newFolder = bpForm.Items.Add("GedFolder", BoFormItemTypes.it_FOLDER);
                newFolder.Width = lastFolder.Width;
                newFolder.Height = lastFolder.Height;
                newFolder.Top = lastFolder.Top;
                newFolder.Left = lastFolder.Left + lastFolder.Width;
                newFolder.Visible = true;
                SAPbouiCOM.Folder gedFolder = ((SAPbouiCOM.Folder)(newFolder.Specific));
                gedFolder.Caption = "GED";
                gedFolder.GroupWith("9");

                // Cria os controles pertencentes a aba GED
                searchPanel = new SearchPanel(bpForm, GED_TAB, newFolder.Top + 25, 25, bpForm.ClientWidth - 50, bpForm.ClientHeight);
                searchPanel.CreateControls(gedContext.SearchFields);

                String uploadPressed = FileResource.MapDesktopResource(@"Images\UploadPressed.png");
                String uploadReleased = FileResource.MapDesktopResource(@"Images\UploadReleased.png");
                String uploadDisabled = FileResource.MapDesktopResource(@"Images\UploadDisabled.png");
                uploadButton = new SboFlatButton(bpForm, GED_TAB, "imgUpload", newFolder.Top + 65, bpForm.ClientWidth - 140);
                uploadButton.SetButtonImages(uploadPressed, uploadReleased, uploadDisabled);
                uploadButton.SetState(false);

                String searchPressed = FileResource.MapDesktopResource(@"Images\SearchPressed.png");
                String searchReleased = FileResource.MapDesktopResource(@"Images\SearchReleased.png");
                String searchDisabled = FileResource.MapDesktopResource(@"Images\SearchDisabled.png");
                searchButton = new SboFlatButton(bpForm, GED_TAB, "imgSearch", newFolder.Top + 65, bpForm.ClientWidth - 100);
                searchButton.SetButtonImages(searchPressed, searchReleased, searchDisabled);
                searchButton.SetState(false);

                SAPbouiCOM.Item docGrid = bpForm.Items.Add("gridDocs", BoFormItemTypes.it_GRID);
                docGrid.FromPane = GED_TAB;
                docGrid.ToPane = GED_TAB;
                docGrid.Left = 25;
                docGrid.Top = newFolder.Top + searchPanel.Height + 35;
                docGrid.Width = searchPanel.Width + 16;
                docGrid.Height = searchPanel.Height * 2;
                docGrid.Visible = false;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void ResizeGedTab(String formUID)
        {
            // Reposiciona os controles de acordo com o resize do form realizado pelo usuário
            try
            {
                SAPbouiCOM.Form bpForm = GetSBOForm(formUID);
                SAPbouiCOM.Item searchPanel = bpForm.Items.Item("pnlSearch");
                searchPanel.Width = bpForm.ClientWidth - 50;
                SAPbouiCOM.Item uploadButton = bpForm.Items.Item("imgUpload");
                uploadButton.Left = bpForm.ClientWidth - 140;
                SAPbouiCOM.Item searchButton = bpForm.Items.Item("imgSearch");
                searchButton.Left = bpForm.ClientWidth - 100;
                SAPbouiCOM.Item gridItem = bpForm.Items.Item("gridDocs");
                gridItem.Width = searchPanel.Width + 16;
                gridItem.Height = searchPanel.Height * 2;
                SAPbouiCOM.Grid gridSpecific = (SAPbouiCOM.Grid)gridItem.Specific;
                gridSpecific.Columns.Item(0).Width = 200;
                gridSpecific.Columns.Item(1).Width = 300;
                gridSpecific.Columns.Item(2).Width = 40;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private void OpenGedTab(String formUID)
        {
            // Altera o pane level do form para GED_TAB, fazendo com que o SAP exiba apenas controles desta aba
            SAPbouiCOM.Form bpForm = GetSBOForm(formUID);
            if (bpForm != null) bpForm.PaneLevel = GED_TAB;
        }

        private void CardCodeChanged(String formUID, BoEventTypes eventType)
        {
            SAPbouiCOM.Form targetForm = GetSBOForm(formUID);
            if (targetForm == null) return;

            SAPbouiCOM.Item cardCodeItem = targetForm.Items.Item("5");
            SAPbouiCOM.EditText cardCodeSpecific = (SAPbouiCOM.EditText)cardCodeItem.Specific;
            // Desabilita os botões caso o código do cliente esteja vazio
            if (String.IsNullOrEmpty(cardCodeSpecific.Value))
            {
                uploadButton.SetState(false);
                searchButton.SetState(false);
            }
        }

        private void UploadButtonClicked(String formUID)
        {
            // Aborta caso o botão esteja desabilitado
            if (uploadButton.Enabled == false) return;

            // Evita duplos cliques no botão de upload
            if (uploadButton.Pressed == true) return;

            // Deixa o botão em estado pressinado
            uploadButton.Press();

            // Abre o dialogo de seleção de arquivo, solta o botão ao termino do upload
            SboOpenFileDialog openFileDialog = new SboOpenFileDialog(sboApplication, formUID, "*.PDF");
            openFileDialog.ShowDialog();
        }

        private void SearchButtonClicked(String formUID)
        {
            // Aborta caso o botão esteja desabilitado
            if (searchButton.Enabled == false) return;

            // Evita duplos cliques no botão de busca
            if (searchButton.Pressed == true) return;

            // Deixa o botão em estado pressinado
            searchButton.Press();

            // Inicia a busca de maneira assíncrona, solta o botão ao termino da execução da Thread
            gedContext.BeginSearch(null, searchPanel.GetSearchKeys(), formUID + ".gridDocs");
        }

        private void SelectCell(String columnName, int row, String formUID)
        {
            if (row == -1) return;
            SAPbouiCOM.Form sboForm = GetSBOForm(formUID);
            if (sboForm == null) return;

            SAPbouiCOM.Item gridItem = sboForm.Items.Item("gridDocs");
            SAPbouiCOM.Grid gridSpecific = ((SAPbouiCOM.Grid)(gridItem.Specific));
            gridSpecific.Rows.SelectedRows.Clear();
            gridSpecific.Rows.SelectedRows.Add(row);
            // Verifica se a coluna com o icone de download foi pressionada
            if (columnName == "open")
            {
                String contentName = (String)gridSpecific.DataTable.GetValue("docName", row);
                String contentRelativeUrl = (String)gridSpecific.DataTable.GetValue("fileUrl", row);
                String contentId = Cipher.GenerateHash(contentName + contentRelativeUrl);

                DocLibContent content = gedContext.ContentDictionary[contentId];
                if (content.ContentType == FileSystemObjectType.Folder)
                    gedContext.BeginSearch(contentRelativeUrl, null, formUID + ".gridDocs");
                if (content.ContentType == FileSystemObjectType.File)
                    gedContext.BeginRetrieve(content);
            }
        }

        public void FileSelected(String formUID)
        {
            // Obtem caminho e nome do arquivo
            SAPbouiCOM.Form openFileDialog = GetSBOForm(formUID);
            SAPbouiCOM.Item txtPath = openFileDialog.Items.Item("txtPath");
            String filePath = ((SAPbouiCOM.EditText)txtPath.Specific).Value;
            SAPbouiCOM.Item txtFile = openFileDialog.Items.Item("txtFile");
            String fileName = ((SAPbouiCOM.EditText)txtFile.Specific).Value;
            // Obtem o form de onde a seleção de arquivo foi iniciada
            UserDataSource parentFormDs = openFileDialog.DataSources.UserDataSources.Item("parentForm");
            String parentForm = parentFormDs.Value;
            // Fecha o dialogo de escolha de arquivo
            openFileDialog.Close();

            // Inicia o upload do documento de maneira assíncrona, solta o botão ao termino da execução da Thread
            SAPbouiCOM.Form bpForm = GetSBOForm(parentForm);
            if (bpForm == null) return;
            SAPbouiCOM.Item cardCodeItem = bpForm.Items.Item("5");
            SAPbouiCOM.EditText cardCodeSpecific = (SAPbouiCOM.EditText)cardCodeItem.Specific;
            NameValueCollection metadata = new NameValueCollection();
            metadata.Add("Cliente", cardCodeSpecific.Value);
            DocLibFileUpload fileUploadData = new DocLibFileUpload(Path.Combine(filePath, fileName), metadata);
            gedContext.BeginStorage(fileUploadData);
        }

        public void FileNotSelected(String formUID)
        {
            // Fecha o dialogo de escolha de arquivo
            SAPbouiCOM.Form openFileDialog = GetSBOForm(formUID);
            openFileDialog.Close();
            // Solta o botão de upload
            if (uploadButton != null) uploadButton.Release();
        }

        public void NotifyEvent(Object eventObject)
        {
            // Retira o botão "Upload" do estado pressionado ao termino do processamento ou em caso de falha
            if (eventObject is UploadFailedEvent)
                uploadButton.Release();
            if (eventObject is UploadFinishedEvent)
                uploadButton.Release();

            // Retira o botão "Search" do estado pressionado ao termino do processamento ou em caso de falha
            if (eventObject is SearchFailedEvent)
                searchButton.Release();
            if (eventObject is SearchFinishedEvent)
                searchButton.Release();
        }

        public Object ShowInfo(String information)
        {
            if (sboApplication == null) return null;

            int dialogWidth = information.Length * 6;
            if (dialogWidth < 350) dialogWidth = 350;

            // Gera um identificador único para o dialogo e verifica se foi aberto mais de uma vez
            String dialogUID = Cipher.GenerateHash(DateTime.Now.ToString("HH:mm:ss.fff"));
            SAPbouiCOM.Form infoDialog = GetSBOForm(dialogUID);
            if (infoDialog != null) return infoDialog; // Retorna o dialogo que já está aberto

            FormCreationParams formParams = (FormCreationParams)(sboApplication.CreateObject(BoCreatableObjectType.cot_FormCreationParams));
            formParams.BorderStyle = BoFormBorderStyle.fbs_Fixed;
            formParams.FormType = "ShowInfo";
            formParams.UniqueID = dialogUID;

            infoDialog = sboApplication.Forms.AddEx(formParams);
            infoDialog.AutoManaged = false;
            infoDialog.Title = "Information";
            infoDialog.Top = 150;
            infoDialog.Left = 330;
            infoDialog.ClientWidth = dialogWidth;
            infoDialog.ClientHeight = 100;

            SAPbouiCOM.Item labelItem = infoDialog.Items.Add("lblInfo", BoFormItemTypes.it_STATIC);
            labelItem.Left = 10;
            labelItem.Top = 50;
            labelItem.Width = dialogWidth;
            labelItem.Height = 25;
            labelItem.AffectsFormMode = false;
            SAPbouiCOM.StaticText labelSpecific = ((SAPbouiCOM.StaticText)(labelItem.Specific));
            labelSpecific.Caption = information;

            infoDialog.Visible = true;
            return infoDialog;
        }

        public void CloseInfo(Object infoDialog)
        {
            if (infoDialog == null) return;

            if (infoDialog is SAPbouiCOM.Form)
                ((SAPbouiCOM.Form)infoDialog).Close();
        }

        private void ClearGrid(String gridOwner, String gridName)
        {
            SAPbouiCOM.Form ownerForm = GetSBOForm(gridOwner);
            if (ownerForm == null) return;

            SAPbouiCOM.Item gridItem = ownerForm.Items.Item(gridName);
            SAPbouiCOM.Grid gridSpecific = ((SAPbouiCOM.Grid)(gridItem.Specific));

            SAPbouiCOM.DataTable gridTable = GetSBODataTable(ownerForm, "gridTable");
            if (gridTable == null) gridTable = ownerForm.DataSources.DataTables.Add("gridTable");
            gridSpecific.DataTable = gridTable;
            gridSpecific.DataTable.Clear();
        }

        public Object PrepareGrid(String gridOwner, String gridName, List<Object> columns, int rowCount)
        {
            SAPbouiCOM.Form ownerForm = GetSBOForm(gridOwner);
            if (ownerForm == null) return null;

            SAPbouiCOM.Item gridItem = ownerForm.Items.Item(gridName);
            gridItem.Visible = true;
            SAPbouiCOM.Grid gridSpecific = ((SAPbouiCOM.Grid)(gridItem.Specific));

            try
            {
                // Monta o Datatable para preencimento do grid
                SAPbouiCOM.DataTable gridTable = GetSBODataTable(ownerForm, "gridTable");
                if (gridTable == null) gridTable = ownerForm.DataSources.DataTables.Add("gridTable");
                gridSpecific.DataTable = gridTable;
                gridSpecific.DataTable.Clear();
                foreach (Object column in columns)
                {
                    Object[] columnProperties = (Object[])column;
                    gridSpecific.DataTable.Columns.Add((String)columnProperties[0], BoFieldsType.ft_Text, 200);
                    int columnIndex = gridSpecific.DataTable.Columns.Count - 1;
                    gridSpecific.Columns.Item(columnIndex).Editable = false;
                    gridSpecific.Columns.Item(columnIndex).TitleObject.Caption = (String)columnProperties[1];
                    gridSpecific.Columns.Item(columnIndex).Width = (int)columnProperties[2];
                    String columnType = (String)columnProperties[3];
                    if ((!String.IsNullOrEmpty(columnType)) && (columnType == "Picture"))
                        gridSpecific.Columns.Item(columnIndex).Type = BoGridColumnType.gct_Picture;
                }
                gridSpecific.DataTable.Rows.Add(rowCount);
                gridSpecific.RowHeaders.Width = 30;
                gridSpecific.SelectionMode = BoMatrixSelect.ms_Auto;
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return null;
            }

            return gridItem;
        }

        public void AddGridRow(Object grid, int rowIndex, Object[] values)
        {
            SAPbouiCOM.Item gridItem = null;
            if (grid is SAPbouiCOM.Item) gridItem = (SAPbouiCOM.Item)grid;
            if (gridItem == null) return;

            try
            {
                SAPbouiCOM.Grid gridSpecific = ((SAPbouiCOM.Grid)(gridItem.Specific));
                int columnIndex = 0;
                foreach (Object value in values)
                {
                    String columnName = gridSpecific.Columns.Item(columnIndex).UniqueID;
                    BoGridColumnType columnType = gridSpecific.Columns.Item(columnIndex).Type;
                    if (columnType == BoGridColumnType.gct_EditText)
                        gridSpecific.DataTable.SetValue(columnName, rowIndex, (String)value);
                    if (columnType == BoGridColumnType.gct_Picture)
                    {
                        PictureColumn column = (PictureColumn)gridSpecific.Columns.Item(columnIndex);
                        FileSystemObjectType columnValue = (FileSystemObjectType)value;
                        if (columnValue == FileSystemObjectType.Folder)
                            column.SetPath(rowIndex, FileResource.MapDesktopResource(@"Images\Folder.png"));
                        if (columnValue == FileSystemObjectType.File)
                            column.SetPath(rowIndex, FileResource.MapDesktopResource(@"Images\Download.png"));
                    }
                    columnIndex++;
                }
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private SAPbouiCOM.Form GetSBOForm(String formName)
        {
            SAPbouiCOM.Form sboForm = null;
            try
            {
                sboForm = sboApplication.Forms.Item(formName);
            }
            catch
            {
                return null;
            }
            return sboForm;
        }

        private SAPbouiCOM.DataTable GetSBODataTable(SAPbouiCOM.Form sboForm, String tableName)
        {
            if (sboForm == null) return null;

            SAPbouiCOM.DataTable sboDataTable = null;
            try
            {
                sboDataTable = sboForm.DataSources.DataTables.Item(tableName);
            }
            catch
            {
                return null;
            }
            return sboDataTable;
        }
    }

}
