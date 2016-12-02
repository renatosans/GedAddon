using System;
using System.Collections.Specialized;


namespace SystemIntegration
{
    public class DocLibFileUpload
    {
        private String filePath;

        private NameValueCollection metadata;


        public String FilePath
        {
            get { return filePath; }
        }

        public NameValueCollection Metadata
        {
            get { return metadata; }
        }

        public DocLibFileUpload(String filePath, NameValueCollection metadata)
        {
            this.filePath = filePath;
            this.metadata = metadata;
        }
    }

}
