using System;
using System.IO;
using System.Collections.Generic;
using System.Collections.Specialized;
using DocMageFramework.AppUtils;
using Microsoft.SharePoint.Client;
using SharepointFile = Microsoft.SharePoint.Client.File;


namespace SystemIntegration
{
    public class DocLibManager
    {
        private ClientContext clientContext;

        private String libraryName;

        private List<DocLibField> fields;

        private List<DocLibContent> contents;

        public List<DocLibField> Fields
        {
            get { return fields; }
        }

        public List<DocLibContent> Contents
        {
            get { return contents; }
        }


        public DocLibManager(String siteUrl, String libraryName, String username, String password)
        {
            clientContext = new ClientContext(siteUrl);
            clientContext.AuthenticationMode = ClientAuthenticationMode.FormsAuthentication;
            clientContext.FormsAuthenticationLoginInfo = new FormsAuthenticationLoginInfo(username, password);
            //clientContext.Credentials = new NetworkCredential(username, password);
            this.libraryName = libraryName;
        }

        public Boolean RetrieveLibraryFields()
        {
            lastError = null;
            fields = new List<DocLibField>();

            List list = null;
            FieldCollection listFields = null;
            try
            {
                list = clientContext.Web.Lists.GetByTitle(libraryName);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                listFields = list.Fields;
                clientContext.Load(listFields);
                clientContext.ExecuteQuery();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return false;
            }

            // Preenche a lista com os campos retornados pela query
            foreach (Field spField in listFields)
            {
                String[] fieldValues = new String[] { spField.DefaultValue }; // Valor default
                if (spField.FieldTypeKind == FieldType.Choice)
                {
                    FieldChoice spChoice = (FieldChoice)spField;
                    fieldValues = new String[spChoice.Choices.Length];
                    spChoice.Choices.CopyTo(fieldValues, 0);
                }

                Boolean allowedType = (spField.FieldTypeKind == FieldType.Text);
                allowedType = allowedType || (spField.FieldTypeKind == FieldType.Choice);
                allowedType = allowedType || (spField.FieldTypeKind == FieldType.DateTime);

                if ((spField.Hidden == false) && (allowedType))
                {
                    DocLibField field = new DocLibField(spField.Title, spField.FieldTypeKind, fieldValues, spField.InternalName);
                    fields.Add(field);
                }
            }

            return true;
        }

        // Convenções:   - As chaves de busca são compostas por "nome" e "valores"
        //               - Os "valores" devem ser um array de Strings com 1 ou 2 elementos (outras quantidades
        //                 de elementos não devem ser passadas)
        private String BuildQuery(Dictionary<String, String[]> searchKeys)
        {
            String queryFormat = @"<View Scope='Recursive'><Query><Where>{0}</Where></Query></View>";
            String camlQuery = "";
            int keyCount = 0;

            foreach (String key in searchKeys.Keys)
            {
                String searchKey = key;
                String[] searchValues = searchKeys[key];
                String expressionOperator = "And";

                if (keyCount > 0) camlQuery = "<" + expressionOperator + ">" + camlQuery;
                if (searchValues.Length == 1) // verificar se o valor está contido
                {
                    camlQuery += "<Contains><FieldRef Name='" + searchKey + "' /><Value Type='Text'>" + searchValues[0] + "</Value></Contains>";
                }
                if (searchValues.Length == 2) // faixa de valores (trazer valores entre o primeiro e o segundo)
                {
                    camlQuery += "<And><Geq><FieldRef Name='" + searchKey + "' /><Value Type='Text'>" + searchValues[0] + "</Value></Geq>";
                    camlQuery += "<Leq><FieldRef Name='" + searchKey + "' /><Value Type='Text'>" + searchValues[1] + "</Value></Leq></And>";
                }
                if (keyCount > 0) camlQuery = camlQuery + "</" + expressionOperator + ">";
                keyCount++;
            }

            return String.Format(queryFormat, camlQuery);
        }

        public Boolean RetrieveLibraryContents(String folderRelativeUrl, Dictionary<String, String[]> searchKeys)
        {
            lastError = null;
            contents = new List<DocLibContent>();

            CamlQuery camlQuery = new CamlQuery();
            if (!String.IsNullOrEmpty(folderRelativeUrl))
                camlQuery.FolderServerRelativeUrl = folderRelativeUrl;
            camlQuery.ViewXml = @"<View><Query><Where></Where></Query></View>";
            if ((searchKeys != null) && (searchKeys.Count > 0))
                camlQuery.ViewXml = BuildQuery(searchKeys);

            List list = clientContext.Web.Lists.GetByTitle(libraryName);
            ListItemCollection listItems = list.GetItems(camlQuery);
            try
            {
                clientContext.Load(list);
                clientContext.Load(listItems);
                clientContext.ExecuteQuery();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return false;
            }

            // Preenche a lista com os documentos retornados pela query
            foreach (ListItem item in listItems)
            {
                DocLibContent content = new DocLibContent((String)item["FileLeafRef"], (String)item["FileRef"], item.FileSystemObjectType);
                contents.Add(content);
            }

            return true;
        }

        public Boolean StoreFile(String fileName, Stream fileContent, NameValueCollection metadata)
        {
            Byte[] buffer;
            List list;
            Folder destinationFolder;
            try
            {
                buffer = new Byte[fileContent.Length];
                fileContent.Read(buffer, 0, buffer.Length);

                list = clientContext.Web.Lists.GetByTitle(libraryName);
                destinationFolder = list.RootFolder;
                clientContext.Load(list);
                clientContext.Load(destinationFolder);
                clientContext.ExecuteQuery();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return false;
            }

            try
            {
                FileCreationInformation fileParams = new FileCreationInformation();
                fileParams.Url = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                fileParams.Content = buffer;

                SharepointFile docLibFile = list.RootFolder.Files.Add(fileParams);
                clientContext.Load(docLibFile);
                clientContext.ExecuteQuery();

                ListItem metadataFields = docLibFile.ListItemAllFields;
                foreach(String entry in metadata)
                {
                    String fieldName = entry;
                    String fieldValue = metadata[entry];
                    metadataFields[fieldName] = fieldValue;
                    metadataFields.Update();
                }
                clientContext.ExecuteQuery();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                return false;
            }

            return true;
        }

        public String RetrieveFile(String fileRelativeUrl, String outputDir)
        {
            lastError = null;
            String outputFile = Path.Combine(outputDir, Path.GetFileName(fileRelativeUrl));
            FileStream fileStream = new FileStream(outputFile, FileMode.Create);
            try
            {
                FileInformation fileInformation = SharepointFile.OpenBinaryDirect(clientContext, fileRelativeUrl);
                IOHandler.CopyStream(fileInformation.Stream, fileStream);
                fileStream.Flush();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
                outputFile = null;
            }
            finally
            {
                fileStream.Close();
            }
            return outputFile;
        }

        private String lastError;

        public String LastError
        {
            get { return lastError; }
        }
    }

}
