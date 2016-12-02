using System;
using System.IO;
using System.Xml;
using System.Collections.Generic;


namespace SystemIntegration
{
    public class GedSearchKeys
    {
        private String xmlFilename;

        private List<String> keyNames;

        public List<String> KeyNames
        {
            get { return keyNames; }
        }


        public GedSearchKeys()
        {
            String dataDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            String gedDataDirectory = Path.Combine(dataDirectory, "GedAddon");
            if (!Directory.Exists(gedDataDirectory)) Directory.CreateDirectory(gedDataDirectory);

            this.xmlFilename = Path.Combine(gedDataDirectory, "SearchKeys.xml");
            this.keyNames = new List<String>();
        }

        private Boolean MountXmlFile()
        {
            // aborta pois o arquivo já foi montado
            if (File.Exists(xmlFilename)) return true;

            List<String> defaultValues = new List<String>();
            defaultValues.Add("Tipo_x0020_Documento");
            defaultValues.Add("Created");
            SaveToXml(defaultValues);

            // falha ao montar o arquivo
            if (lastError != null) return false;

            return true;
        }

        public void LoadFromXml()
        {
            lastError = null;

            // Sai do método caso o arquivo de configurações não possa ser utilizado
            if (!MountXmlFile()) return;

            try
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(xmlFilename);
                XmlNode mainNode = xmlDoc.SelectSingleNode("//searchkeys");
                foreach (XmlNode childNode in mainNode.ChildNodes)
                {
                    // Adiciona o valor do nó(xml) a lista de chaves
                    keyNames.Add(childNode.InnerText);
                }
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        public void SaveToXml(List<String> keyNames)
        {
            lastError = null;

            try
            {
                StreamWriter streamWriter = File.CreateText(xmlFilename);
                streamWriter.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
                streamWriter.WriteLine("<searchkeys>");
                foreach (String name in keyNames)
                    streamWriter.WriteLine("  <key>" + name + "</key>");
                streamWriter.WriteLine("</searchkeys>");
                streamWriter.Flush();
                streamWriter.Close();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        private String lastError;

        public String LastError
        {
            get { return lastError; }
        }
    }

}
