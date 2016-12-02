using System;


namespace SystemIntegration
{
    public class SearchFailedEvent
    {
        private String reason;

        public String Reason
        {
            get { return reason; }
        }


        public SearchFailedEvent(String reason)
        {
            this.reason = reason;
        }
    }

}
