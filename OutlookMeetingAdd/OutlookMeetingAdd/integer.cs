using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookMeetingAdd
{
    class integer
    {
        public string MeetingTime;
        public string MeetingRoomLocation;
        public string Floor;

        public integer()
        {
            MeetingTime = null;
            MeetingRoomLocation = null;
            Floor = null;
        }


        public integer(string MeetingTime,string MeetingRoomLocation,string Floor)
        {
            this.MeetingTime = MeetingTime;
            this.MeetingRoomLocation = MeetingRoomLocation;
            this.Floor = Floor;
        }
    }
    
}
