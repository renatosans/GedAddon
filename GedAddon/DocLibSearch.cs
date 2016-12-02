using System;
using System.Collections.Generic;


namespace SystemIntegration
{
    public class DocLibSearch
    {
        private String folderRelativeUrl;

        private Dictionary<String, String[]> searchKeys;

        private String display;

        public String FolderRelativeUrl
        {
            get { return folderRelativeUrl; }
        }

        public Dictionary<String, String[]> SearchKeys
        {
            get { return searchKeys; }
        }

        public String Display
        {
            get { return display; }
        }


        public DocLibSearch(String folderRelativeUrl, Dictionary<String, String[]> searchKeys, String display)
        {
            this.folderRelativeUrl = folderRelativeUrl;
            this.searchKeys = searchKeys;
            this.display = display;
        }
    }

}
