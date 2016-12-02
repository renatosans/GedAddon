using System;
using System.IO;
using System.Xml;


namespace SystemIntegration
{
    public class GedSettings
    {
        private String xmlFilename;

        private String server;
        private String library;
        private String username;
        private String password;

        public String Server
        {
            get { return server; }
        }

        public String Library
        {
            get { return library; }
        }

        public String Username
        {
            get { return username; }
        }

        public String Password
        {
            get { return password; }
        }


        public GedSettings()
        {
            String dataDirectory = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            String gedDataDirectory = Path.Combine(dataDirectory, "GedAddon");
            if (!Directory.Exists(gedDataDirectory)) Directory.CreateDirectory(gedDataDirectory);

            this.xmlFilename = Path.Combine(gedDataDirectory, "Settings.xml");
        }

        private Boolean MountXmlFile()
        {
            // aborta pois o arquivo já foi montado
            if (File.Exists(xmlFilename)) return true;

            // Monta o XML com os valores default
            String defaultServer = "https://www.dataged.com.br";
            String defaultDocLib = "Datacopy - Comercial";
            SaveToXml(defaultServer, defaultDocLib, "", "");

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
                XmlTextReader xmlReader = new XmlTextReader(xmlFilename);
                xmlReader.ReadStartElement("settings");
                server = xmlReader.ReadElementString("server");
                library = xmlReader.ReadElementString("library");
                username = xmlReader.ReadElementString("username");
                password = xmlReader.ReadElementString("password");
                xmlReader.ReadEndElement();
                xmlReader.Close();
            }
            catch (Exception exc)
            {
                lastError = exc.Message;
            }
        }

        public void SaveToXml(String server, String library, String username, String password)
        {
            lastError = null;

            try
            {
                StreamWriter streamWriter = File.CreateText(xmlFilename);
                streamWriter.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
                streamWriter.WriteLine("<settings>");
                streamWriter.WriteLine("  <server>" + server + "</server>");
                streamWriter.WriteLine("  <library>" + library + "</library>");
                streamWriter.WriteLine("  <username>" + username + "</username>");
                streamWriter.WriteLine("  <password>" + password + "</password>");
                streamWriter.WriteLine("</settings>");
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
