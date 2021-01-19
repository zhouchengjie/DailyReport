using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DailyReport
{
    public class ChannelPercentageRankDO : DataObjectBase
    {
        private int iD;
        public int ID
        {
            get { return iD; }
            set { iD = value; }
        }

        private string channelName;
        public string ChannelName
        {
            get { return channelName; }
            set { channelName = value; }
        }

        private decimal tVPercentage;
        public decimal TVPercentage
        {
            get { return tVPercentage; }
            set { tVPercentage = value; }
        }

        private DateTime effectiveDate;
        public DateTime EffectiveDate
        {
            get { return effectiveDate; }
            set { effectiveDate = value; }
        }

        public ChannelPercentageRankDO()
        {
            this.BO_Name = "ChannelPercentageRank";
            this.PK_Name = "ID";
        }
    }
}
