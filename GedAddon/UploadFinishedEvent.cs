using System;


namespace SystemIntegration
{
    public class UploadFinishedEvent
    {
        private int fileSize;

        public int FileSize
        {
            get { return fileSize; }
        }


        public UploadFinishedEvent(int fileSize)
        {
            this.fileSize = fileSize;
        }
    }

}
