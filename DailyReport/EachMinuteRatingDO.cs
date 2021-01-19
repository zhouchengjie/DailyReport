using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DailyReport
{
    public class EachMinuteRatingDO : DataObjectBase
    {
        private int iD;
        public int ID
        {
            get { return iD; }
            set { iD = value; }
        }

        private string programName;
        public string ProgramName
        {
            get { return programName; }
            set { programName = value; }
        }

        private string strMinute;
        public string StrMinute
        {
            get { return strMinute; }
            set { strMinute = value; }
        }

        private int minuteIndex;
        public int MinuteIndex
        {
            get { return minuteIndex; }
            set { minuteIndex = value; }
        }

        private DateTime effectiveDate;
        public DateTime EffectiveDate
        {
            get { return effectiveDate; }
            set { effectiveDate = value; }
        }


        private decimal tVRate;
        public decimal TVRate
        {
            get { return tVRate; }
            set { tVRate = value; }
        }

        private decimal tVPercentage;
        public decimal TVPercentage
        {
            get { return tVPercentage; }
            set { tVPercentage = value; }
        }

        private decimal tVLoyalty;
        public decimal TVLoyalty
        {
            get { return tVLoyalty; }
            set { tVLoyalty = value; }
        }

        public EachMinuteRatingDO()
        {
            this.BO_Name = "EachMinuteRating";
            this.PK_Name = "ID";
        }
    }
}
