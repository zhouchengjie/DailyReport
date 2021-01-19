using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DailyReport
{
    public class EachDayTVSeriesDO : DataObjectBase
    {
        private int iD;
        public int ID
        {
            get { return iD; }
            set { iD = value; }
        }

        private string strTime;
        public string StrTime
        {
            get { return strTime; }
            set { strTime = value; }
        }

        private int timeTypeID;
        public int TimeTypeID
        {
            get { return timeTypeID; }
            set { timeTypeID = value; }
        }

        private decimal tVRate;
        public decimal TVRate
        {
            get { return tVRate; }
            set { tVRate = value; }
        }

        private DateTime effectiveDate;
        public DateTime EffectiveDate
        {
            get { return effectiveDate; }
            set { effectiveDate = value; }
        }

        public EachDayTVSeriesDO()
        {
            this.BO_Name = "EachDayTVSeries";
            this.PK_Name = "ID";
        }
    }
}
