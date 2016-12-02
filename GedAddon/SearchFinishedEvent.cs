using System;


namespace SystemIntegration
{
    public class SearchFinishedEvent
    {
        private int matches;

        public int Matches
        {
            get { return matches; }
        }


        public SearchFinishedEvent(int matches)
        {
            this.matches = matches;
        }
    }

}
