using System;
using System.IO;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using System.Collections.Specialized;
using DocMageFramework.AppUtils;
using Microsoft.SharePoint.Client;


namespace SystemIntegration
{
    public class GedContext
    {
        private DocLibManager docLibManager;

        private IGraphicalDisplay graphicalDisplay;

        private Dictionary<String, DocLibContent> contentDictionary;

        public Dictionary<String, DocLibContent> ContentDictionary
        {
            get { return contentDictionary; }
        }

        private List<DocLibField> searchFields;

        public List<DocLibField> SearchFields
        {
            get { return searchFields; }
        }

        private NameValueCollection fieldNames;

        public NameValueCollection FieldNames
        {
            get { return fieldNames; }
        }


        public GedContext(IGraphicalDisplay graphicalDisplay)
        {
            this.docLibManager = null; // permanece nulo até a conexão
            this.graphicalDisplay = graphicalDisplay;
            this.contentDictionary = new Dictionary<String, DocLibContent>();
            this.searchFields = new List<DocLibField>();
            this.fieldNames = new NameValueCollection();
        }

        public void Connect()
        {
            // Prepara a conexão com o servidor
            GedSettings settings = new GedSettings();
            settings.LoadFromXml();
            Boolean configMissing = false;
            if (String.IsNullOrEmpty(settings.Server)) configMissing = true;
            if (String.IsNullOrEmpty(settings.Library)) configMissing = true;
            if (configMissing)
            {
                graphicalDisplay.ShowInfo("O servidor de GED não está configurado." + settings.LastError);
                return;
            }
            docLibManager = new DocLibManager(settings.Server, settings.Library, settings.Username, settings.Password);

            // Preenche a lista de campos que devem ser utilizados na pesquisa
            GetSearchFields();
        }

        private void GetSearchFields()
        {
            GedSearchKeys gedSearchKeys = new GedSearchKeys();
            gedSearchKeys.LoadFromXml();
            if (gedSearchKeys.LastError != null)
            {
                // aborta pois não conseguiu ler o arquivo de configuração
                graphicalDisplay.ShowInfo("Falha no GED. " + gedSearchKeys.LastError);
                return;
            }

            // Busca os campos existentes na biblioteca de documentos
            Boolean fieldsRetrieved = docLibManager.RetrieveLibraryFields();
            if (!fieldsRetrieved)
            {
                // aborta pois a requisição não conseguiu recuperar os campos
                graphicalDisplay.ShowInfo("Falha no GED. " + docLibManager.LastError);
                return;
            }

            // Filtra somente os campos especificados na lista "fieldnameFilter"
            List<String> fieldnameFilter = gedSearchKeys.KeyNames;
            List<DocLibField> filteredFields = new List<DocLibField>(); // lista temporária
            foreach (DocLibField field in docLibManager.Fields)
                if (fieldnameFilter.Contains(field.InternalName)) filteredFields.Add(field);

            // Atualiza os campos a partir da lista temporária
            if (filteredFields.Count > 0) searchFields.Clear();
            foreach (DocLibField field in filteredFields)
                searchFields.Add(field);

            // Disponibiliza uma coleção com os nomes dos campos para que o usuário escolha
            // quais ele quer na pesquisa
            fieldNames.Clear();
            foreach (DocLibField field in docLibManager.Fields)
                fieldNames.Add(field.InternalName, field.DisplayName);
        }

        public void BeginStorage(DocLibFileUpload fileUploadData)
        {
            if (docLibManager == null)
            {
                graphicalDisplay.ShowInfo("Não foi possível realizar o upload do documento. Verifique as configurações.");
                return;
            }

            // Realiza o upload em uma thread separada para evitar o congelamento da interface
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(StorageHandler);
            backgroundWorker.RunWorkerAsync(fileUploadData);

            System.Windows.Forms.Application.DoEvents();
        }

        private void StorageHandler(Object sender, DoWorkEventArgs eventArgs)
        {
            DocLibFileUpload fileUploadData = (DocLibFileUpload)eventArgs.Argument;

            String fileName = Path.GetFileName(fileUploadData.FilePath);
            Stream fileContent = new FileStream(fileUploadData.FilePath, FileMode.Open);
            int fileSize = (int)fileContent.Length;
            NameValueCollection metadata = fileUploadData.Metadata;
            try
            {
                Boolean fileStored = docLibManager.StoreFile(fileName, fileContent, metadata);
                if (!fileStored)
                {
                    graphicalDisplay.ShowInfo(docLibManager.LastError);
                    graphicalDisplay.NotifyEvent(new UploadFailedEvent(docLibManager.LastError));
                    return;
                }
            }
            finally
            {
                fileContent.Close();
            }
            graphicalDisplay.NotifyEvent(new UploadFinishedEvent(fileSize));
        }

        public void BeginRetrieve(DocLibContent document)
        {
            if (docLibManager == null)
            {
                graphicalDisplay.ShowInfo("Não foi possível recuperar o documento. Verifique as configurações.");
                return;
            }

            // Realiza o download em uma thread separada para evitar o congelamento da interface
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(RetrieveHandler);
            backgroundWorker.RunWorkerAsync(document);

            System.Windows.Forms.Application.DoEvents();
        }

        private void RetrieveHandler(Object sender, DoWorkEventArgs eventArgs)
        {
            DocLibContent document = (DocLibContent)eventArgs.Argument;
            if (document.ContentType != FileSystemObjectType.File) return;

            String dataDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            String gedDataDirectory = Path.Combine(dataDirectory, "GedAddon");
            if (!Directory.Exists(gedDataDirectory)) Directory.CreateDirectory(gedDataDirectory);
            String downloadDirectory = Path.Combine(gedDataDirectory, "Download");
            if (!Directory.Exists(downloadDirectory)) Directory.CreateDirectory(downloadDirectory);

            Object infoDialog = graphicalDisplay.ShowInfo("Fazendo download do documento " + document.Name + "...");
            String retrievedFile = docLibManager.RetrieveFile(document.RelativeUrl, downloadDirectory);
            graphicalDisplay.CloseInfo(infoDialog);
            if (retrievedFile == null)
            {
                graphicalDisplay.ShowInfo(docLibManager.LastError);
                return;
            }

            if (System.IO.File.Exists(retrievedFile)) Process.Start(retrievedFile);
        }

        public void BeginSearch(String folderRelativeUrl, Dictionary<String, String[]> searchKeys, String display)
        {
            if (docLibManager == null)
            {
                graphicalDisplay.ShowInfo("Não foi possível realizar a pesquisa. Verifique as configurações.");
                return;
            }

            // Realiza o processo em uma thread separada para evitar o congelamento da interface
            BackgroundWorker backgroundWorker = new BackgroundWorker();
            backgroundWorker.DoWork += new DoWorkEventHandler(SearchHandler);
            backgroundWorker.RunWorkerAsync(new DocLibSearch(folderRelativeUrl, searchKeys, display));

            System.Windows.Forms.Application.DoEvents();
        }

        private void SearchHandler(Object sender, DoWorkEventArgs eventArgs)
        {
            DocLibSearch doclibSearch = (DocLibSearch)eventArgs.Argument;

            Object infoDialog = graphicalDisplay.ShowInfo("Solicitando dados ao servidor remoto...");
            Boolean contentsRetrieved = docLibManager.RetrieveLibraryContents(doclibSearch.FolderRelativeUrl, doclibSearch.SearchKeys);
            graphicalDisplay.CloseInfo(infoDialog);
            if (!contentsRetrieved)
            {
                graphicalDisplay.ShowInfo(docLibManager.LastError);
                graphicalDisplay.NotifyEvent(new SearchFailedEvent(docLibManager.LastError));
                return;
            }

            UpdateContentDictionary(docLibManager.Contents);
            DisplayContents(doclibSearch.Display, docLibManager.Contents);
            graphicalDisplay.NotifyEvent(new SearchFinishedEvent(docLibManager.Contents.Count));
        }

        private void UpdateContentDictionary(List<DocLibContent> docLibContents)
        {
            if (docLibContents == null) return;

            contentDictionary.Clear();
            foreach (DocLibContent content in docLibContents)
            {
                String contentId = Cipher.GenerateHash(content.Name + content.RelativeUrl);
                contentDictionary.Add(contentId, content);
            }
        }

        private void DisplayContents(String display, List<DocLibContent> docLibContents)
        {
            if (docLibContents == null) return;
            if (docLibContents.Count == 0) // nenhum documento ou pasta na biblioteca
            {
                graphicalDisplay.ShowInfo("Nenhum documento encontrado.");
                return;
            }

            String[] nameParts = display.Split(new Char[] { '.' });
            if (nameParts.Length != 2) return;
            String gridOwner = nameParts[0];
            String gridName = nameParts[1];

            List<Object> columns = new List<Object>();
            columns.Add(new Object[] { "docName", "Nome do Documento", 200, null });
            columns.Add(new Object[] { "fileUrl", "URL", 300, null });
            columns.Add(new Object[] { "open", "Abrir", 40, "Picture" });
            Object grid = graphicalDisplay.PrepareGrid(gridOwner, gridName, columns, docLibContents.Count);
            int rowIndex = 0;
            foreach (DocLibContent content in docLibContents)
            {
                graphicalDisplay.AddGridRow(grid, rowIndex, new Object[] { content.Name, content.RelativeUrl, content.ContentType });
                rowIndex++;
            }
        }
    }

}
