using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;

namespace DailyReport
{
    public class DRManager : BusinessLogicBase
    {
        private static BusinessLogicBase baseLogic = new BusinessLogicBase();

        public static EachProgramRatingDO GetProgramRatingDOByID(int id)
        {
            EachProgramRatingDO projectDO = new EachProgramRatingDO();
            return (EachProgramRatingDO)baseLogic.Select(projectDO, id);
        }

        public static EachProgramRatingDO GetProgramRatingDOByName(string programName, DateTime effectiveDate, int beginSeconds)
        {
            EachProgramRatingDO projectDO = new EachProgramRatingDO();
            string sql = @"select * from EachProgramRating 
where ProgramName=@programName 
and EffectiveDate=@EffectiveDate
and EndSecond>@BeginSecond-1800
and EndSecond<@BeginSecond+1800
";
            SqlParameter[] pars = new SqlParameter[3];
            pars[0] = new SqlParameter("@programName", programName);
            pars[1] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[2] = new SqlParameter("@BeginSecond", beginSeconds);
            return (EachProgramRatingDO)baseLogic.Select(projectDO, sql, pars);
        }

        public static EachProgramRatingDO GetProgramRatingDOByName(string programName, DateTime effectiveDate)
        {
            EachProgramRatingDO projectDO = new EachProgramRatingDO();
            string sql = @"select * from EachProgramRating 
where ProgramName=@programName 
and EffectiveDate=@EffectiveDate";
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@programName", programName);
            pars[1] = new SqlParameter("@EffectiveDate", effectiveDate);
            DataTable dt = baseLogic.Select(sql, pars);
            if (dt != null && dt.Rows.Count == 1)
            {
                int id = 0;
                int.TryParse(dt.Rows[0]["ID"].ToString(), out id);
                projectDO = GetProgramRatingDOByID(id);
            }
            return projectDO;
        }

        public static DataTable GetProgramRatingsByName(string programName, DateTime effectiveDate)
        {
            string sql = @"select * from EachProgramRating 
where ProgramName=@programName 
and EffectiveDate=@EffectiveDate";
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@programName", programName);
            pars[1] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.Select(sql, pars);
        }

        public static DataTable GetAllMinutesRate(DateTime effectiveDate)
        {
            string spName = "GetAllMinutesRate";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.ExceuteStoredProcedure(spName, pars);
        }

        public static DataTable GetChannelPercentageRank(DateTime effectiveDate)
        {
            string sql = @"declare @YesterdayDate datetime
set @YesterdayDate=dateadd(day,-1,@EffectiveDate)

select ROW_NUMBER() over (order by TVPercentage desc) as RowNum,
case when mm.DisplayName is null then rr.ChannelName else mm.DisplayName end as ChannelName,
CAST(rr.TVPercentage as decimal(38,4)) as TVPercentage
into #current
from [dbo].[ChannelPercentageRank] rr left join [dbo].[ChannelNameMapping] mm on rr.ChannelName=mm.ColumnName
where effectivedate=@EffectiveDate
order by tvPercentage desc

select ROW_NUMBER() over (order by TVPercentage desc) as RowNum,
case when mm.DisplayName is null then rr.ChannelName else mm.DisplayName end as ChannelName,
CAST(rr.TVPercentage as decimal(38,4)) as TVPercentage
into #lastday
from [dbo].[ChannelPercentageRank] rr left join [dbo].[ChannelNameMapping] mm on rr.ChannelName=mm.ColumnName
where effectivedate=@YesterdayDate
order by tvPercentage desc

select top 15 cc.*,
case when dd.TVPercentage>0 then (cc.TVPercentage-dd.TVPercentage)*100/dd.TVPercentage
when dd.TVPercentage=0 then 0 else 0 end as changeRate,
cc.RowNum-dd.RowNum as changeRank
from #current cc left join #lastday dd on dd.ChannelName=cc.ChannelName
order by cc.RowNum

drop table #current
drop table #lastday";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.Select(sql, pars);
        }

        public static DataTable GetProgramTargetRateRank(DateTime effectiveDate)
        {
            string spName = "GetProgramTargetRateRank";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.ExceuteStoredProcedure(spName, pars);
        }


        public static void DeleteProgramRatingDO(int id)
        {
            EachProgramRatingDO apo = new EachProgramRatingDO();
            apo.ID = id;
            baseLogic.Delete(apo);
        }

        public static DataTable GetMorningPrograms(DateTime effectiveDate) 
        {
            string sql = @"select * from [dbo].[EachProgramRating] 
where EffectiveDate=@EffectiveDate
and Beginsecond>8.5*3600 and beginsecond<12*3600
and (EndSecond-BeginSecond)>=12*60
and CategoryName<>'新闻/时事'
order by Beginsecond";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.Select(sql, pars);
        }

        public static DataTable GetAfternoonPrograms(DateTime effectiveDate)
        {
            string sql = @"select * from [dbo].[EachProgramRating] 
where EffectiveDate=@EffectiveDate
and Beginsecond>12.5*3600 
and Beginsecond<18*3600
and (EndSecond-BeginSecond)>=12*60
and CategoryName<>'新闻/时事'
order by Beginsecond";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.Select(sql, pars);
        }

        public static DataTable GetEvenningPrograms(DateTime effectiveDate)
        {
            string sql = @"select * from [dbo].[EachProgramRating] 
where EffectiveDate=@EffectiveDate
and Beginsecond>19.5*3600 
and Beginsecond<23*3600
and (EndSecond-BeginSecond)>=12*60
and CategoryName<>'新闻/时事'
order by Beginsecond";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return baseLogic.Select(sql, pars);
        }

        public static EachDayTVSeriesDO GetTVSeries(DateTime effectiveDate, int timeTypeID)
        {
            string sql = @"select * from [dbo].[EachDayTVSeries] 
where EffectiveDate=@EffectiveDate
and TimeTypeID=@TimeTypeID";
            EachDayTVSeriesDO tvDO = new EachDayTVSeriesDO();
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[1] = new SqlParameter("@TimeTypeID", timeTypeID);
            return (EachDayTVSeriesDO)baseLogic.Select(tvDO, sql, pars);
        }

        public static decimal GetLastSameTVSeriesRate(DateTime effectiveDate, int timeTypeID)
        {
            string sql = @"select top 1 TVRate from [dbo].[EachDayTVSeries] 
where EffectiveDate<@EffectiveDate
and EffectiveDate>=dateadd(day,-7,@EffectiveDate)
and TimeTypeID=@TimeTypeID
order by EffectiveDate desc";
            EachDayTVSeriesDO tvDO = new EachDayTVSeriesDO();
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[1] = new SqlParameter("@TimeTypeID", timeTypeID);
            decimal result = 0;
            DataTable dt = baseLogic.Select(sql, pars);
            if (dt != null && dt.Rows.Count > 0)
            {
                decimal.TryParse(dt.Rows[0][0].ToString(), out result);
            }
            return result;
        }

        public static decimal GetLastSameProgramRate(DateTime effectiveDate, string programName, int beginSecond)
        {
            string sql = @"select top 1 TVRate from [EachProgramRating] 
where ProgramName=@ProgramName 
and EffectiveDate<@EffectiveDate 
and EffectiveDate>=dateadd(day,-7,@EffectiveDate)
and BeginSecond>@BeginSecond-1800
and BeginSecond<@BeginSecond+1800
order by EffectiveDate desc";
            SqlParameter[] pars = new SqlParameter[3];
            pars[0] = new SqlParameter("@ProgramName", programName);
            pars[1] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[2] = new SqlParameter("@BeginSecond", beginSecond);

            decimal result = 0;
            DataTable dt = baseLogic.Select(sql, pars);
            if (dt != null && dt.Rows.Count > 0) 
            {
                decimal.TryParse(dt.Rows[0][0].ToString(), out result);
            }
            return result;
        }

        public static decimal GetLastSameProgramRateInSameDay(DateTime effectiveDate, string programName, int beginSecond)
        {
            string sql = @"select top 1 TVRate from [EachProgramRating] 
where ProgramName=@ProgramName 
and EffectiveDate=dateadd(day,-7,@EffectiveDate)
and BeginSecond>@BeginSecond-1800
and BeginSecond<@BeginSecond+1800
order by EffectiveDate desc";
            SqlParameter[] pars = new SqlParameter[3];
            pars[0] = new SqlParameter("@ProgramName", programName);
            pars[1] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[2] = new SqlParameter("@BeginSecond", beginSecond);

            decimal result = 0;
            DataTable dt = baseLogic.Select(sql, pars);
            if (dt != null && dt.Rows.Count > 0)
            {
                decimal.TryParse(dt.Rows[0][0].ToString(), out result);
            }
            return result;
        }

        public static DataTable GetWeekendPrograms(DateTime fromDate, DateTime toDate) 
        {
            string sql = @"select * from [dbo].[EachWeekEndProgramRate] 
where EffectiveDate>=@FromDate 
and EffectiveDate<=@ToDate 
and BeginSecond>=20*3600 
and BeginSecond<22*3600";
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@FromDate", fromDate);
            pars[1] = new SqlParameter("@ToDate", toDate);
            return BusinessLogicBase.Default.Select(sql, pars);
        }

        public static decimal GetWeekendProgramsByName(DateTime effectiveDate, string programName)
        {
            string sql = @"select TVRate from [dbo].[EachWeekEndProgramRate] 
where EffectiveDate=@EffectiveDate 
and ProgramName=@ProgramName
and BeginSecond>=20*3600
and BeginSecond<22*3600";
            SqlParameter[] pars = new SqlParameter[2];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            pars[1] = new SqlParameter("@ProgramName", programName);
            decimal result = 0;
            DataTable dt = baseLogic.Select(sql, pars);
            if (dt != null && dt.Rows.Count > 0)
            {
                decimal.TryParse(dt.Rows[0][0].ToString(), out result);
            }
            return result;
        }

        public static void ClearDataByDate(DateTime effectiveDate)
        {
            string sql = @"delete from [dbo].[ChannelPercentageRank] where EffectiveDate=@EffectiveDate
delete from [dbo].[EachDayTVSeries] where EffectiveDate=@EffectiveDate
delete from [dbo].[EachMinuteRating] where EffectiveDate=@EffectiveDate
delete from [dbo].[EachProgramRating] where EffectiveDate=@EffectiveDate
delete from [dbo].[EachWeekEndProgramRate] where EffectiveDate>=DATEADD(WK, DATEDIFF(WK, 0, @EffectiveDate),-2) and EffectiveDate<=DATEADD(WK, DATEDIFF(WK, 0, @EffectiveDate),-1)";
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            BusinessLogicBase.Default.Execute(sql, pars);
        }

        public static DataTable GetUploadedDataByDate(DateTime effectiveDate, string fileName)
        {
            string sql = string.Empty;

            if (fileName == "每分钟节目")
            {
                sql = "select * from [dbo].[EachProgramRating] where EffectiveDate=@EffectiveDate";
            }
            else if (fileName == "第1动画乐园")
            {
                sql = "select * from [dbo].[EachProgramRating] where EffectiveDate=@EffectiveDate and ProgramName='第1动画乐园'";
            }
            else if (fileName == "三大剧场")
            {
                sql = "select * from [dbo].[EachDayTVSeries] where EffectiveDate=@EffectiveDate";
            }
            else if (fileName == "上星频道市场份额排名")
            {
                sql = "select * from [dbo].[ChannelPercentageRank] where EffectiveDate=@EffectiveDate";
            }
            else if (fileName == "34城市")
            {
                sql = "select * from [dbo].[EachWeekEndProgramRate] where EffectiveDate>=DATEADD(day, -1, @EffectiveDate) and EffectiveDate<=@EffectiveDate";
            }
            SqlParameter[] pars = new SqlParameter[1];
            pars[0] = new SqlParameter("@EffectiveDate", effectiveDate);
            return BusinessLogicBase.Default.Select(sql, pars);
        }
    }
}
