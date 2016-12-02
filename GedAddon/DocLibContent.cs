using System;
using Microsoft.SharePoint.Client;


namespace SystemIntegration
{
    public class DocLibContent
    {
        private String contentName;

        private String contentRelativeUrl;

        private FileSystemObjectType contentType;

        public String Name
        {
            get { return contentName; }
        }

        public String RelativeUrl
        {
            get { return contentRelativeUrl; }
        }

        public FileSystemObjectType ContentType
        {
            get { return contentType; }
        }


        public DocLibContent(String name, String relativeUrl, FileSystemObjectType contentType)
        {
            this.contentName = name;
            this.contentRelativeUrl = relativeUrl;
            this.contentType = contentType;
        }
    }

}
