using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DailyReport
{
    public class EachProgramRatingDO : DataObjectBase
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

        private string channelName;
        public string ChannelName
        {
            get { return channelName; }
            set { channelName = value; }
        }

        private string categoryName;
        public string CategoryName
        {
            get { return categoryName; }
            set { categoryName = value; }
        }

        private string subCategoryName;
        public string SubCategoryName
        {
            get { return subCategoryName; }
            set { subCategoryName = value; }
        }

        private string strBeginTime;
        public string StrBeginTime
        {
            get { return strBeginTime; }
            set { strBeginTime = value; }
        }

        private string strEndTime;
        public string StrEndTime
        {
            get { return strEndTime; }
            set { strEndTime = value; }
        }

        private int beginSecond;
        public int BeginSecond
        {
            get { return beginSecond; }
            set { beginSecond = value; }
        }

        private int endSecond;
        public int EndSecond
        {
            get { return endSecond; }
            set { endSecond = value; }
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

        private DateTime effectiveDate;
        public DateTime EffectiveDate
        {
            get { return effectiveDate; }
            set { effectiveDate = value; }
        }

        public EachProgramRatingDO()
        {
            this.BO_Name = "EachProgramRating";
            this.PK_Name = "ID";
        }
    }
}
