using System;


namespace SystemIntegration
{
    public class UploadFailedEvent
    {
        private String reason;

        public String Reason
        {
            get { return reason; }
        }


        public UploadFailedEvent(String reason)
        {
            this.reason = reason;
        }
    }

}
