using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Configuration;
using Org.BouncyCastle.Asn1.IsisMtt.X509;

namespace DailyReport
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();

            LoadYesteday();
        }

        private void LoadYesteday() 
        {
            DateTime yesDate = DateTime.Today.AddDays(-1);
            this.calReportDate.Value = yesDate;
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (ddlFileType.SelectedIndex < 0)
            {
                MessageBox.Show("请选择读取的文件类型");
                return;
            }
            string fileTypeName = ddlFileType.SelectedItem.ToString();
            if (this.ofdReadFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string filePath = this.ofdReadFile.FileName;
                if (string.IsNullOrEmpty(filePath))
                {
                    return;
                }
                string fileName = Path.GetFileNameWithoutExtension(filePath);
                if (fileTypeName != fileName)
                {
                    MessageBox.Show("请重新选择文件，您选择读取的文件类型是{" + fileTypeName + "}，但是您选择的文件不是该类型文件。");
                    return;
                }
                DateTime reportDate = calReportDate.Value;
                if (fileTypeName == "每分钟节目")
                {
                    ReadEachMinuteProgramFile(reportDate, filePath);
                }
                else if (fileTypeName == "第1动画乐园")
                {
                    ReadSpecialProgramFile(reportDate, filePath);
                }
                else if (fileTypeName == "三大剧场")
                {
                    ReadTVSeries(reportDate, filePath);
                }
                else if (fileTypeName == "上星频道市场份额排名")
                {
                    ReadChannelRankFile(reportDate, filePath);
                }
                else if (fileTypeName == "34城市")
                {
                    Read34CitiesFile(reportDate, filePath);
                }

                LoadData();
            }
        }

        private void Read34CitiesFile(DateTime reportDate, string filePath)
        {
            try
            {
                DataTable dtChannel = NPOIHelper.ImportExceltoDt(filePath, "节目", 1);
                if (dtChannel == null || dtChannel.Rows.Count <= 0)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                for (int i = 1; i < dtChannel.Rows.Count; i++)
                {
                    DataRow row = dtChannel.Rows[i];
                    string programName = row[0].ToString().Trim();
                    if (string.IsNullOrEmpty(programName))
                    {
                        continue;
                    }
                    
                    string channelName = row[1].ToString().Trim();
                    string strDate = row[2].ToString().Trim();
                    string weekday = row[3].ToString().Trim();
                    string strBeginTime = row[4].ToString().Trim();
                    string strEndTime = row[6].ToString().Trim();
                    string categoryName = row[7].ToString().Trim();
                    string subCategoryName = row[8].ToString().Trim();
                    string tvRate = row[9].ToString();
                    decimal _tvRate = 0;
                    decimal.TryParse(tvRate, out _tvRate);
                    int beginSeconds = CommonFunction.GetTimeToSeconds(strBeginTime);
                    int endSeconds = CommonFunction.GetTimeToSeconds(strEndTime);
                    DateTime effectiveDate = Constants.Date_Min;
                    DateTime.TryParse(strDate, out effectiveDate);

                    EachWeekEndProgramRateDO tvDO = new EachWeekEndProgramRateDO();
                    tvDO.BeginSecond = beginSeconds;
                    tvDO.CategoryName = categoryName;
                    tvDO.ChannelName = channelName;
                    tvDO.EffectiveDate = effectiveDate;
                    tvDO.EndSecond = endSeconds;
                    tvDO.ProgramName = programName;
                    tvDO.StrBeginTime = strBeginTime;
                    tvDO.StrEndTime = strEndTime;
                    tvDO.StrWeekDay = weekday;
                    tvDO.SubCategoryName = subCategoryName;
                    tvDO.TVRate = _tvRate;
                    BusinessLogicBase.Default.Insert(tvDO);
                }

                MessageBox.Show("读取成功！");
            }
            catch (Exception ex)
            {

            }
        }

        private void ReadTVSeries(DateTime reportDate, string filePath)  
        {
            try
            {
                DataTable dtChannel = NPOIHelper.ImportExceltoDt(filePath, "交互分析", 1);
                if (dtChannel == null)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                for (int i = 2; i < dtChannel.Rows.Count; i++)
                {
                    DataRow row = dtChannel.Rows[i];
                    string name = row[1].ToString().Trim();
                    if (string.IsNullOrEmpty(name)) 
                    {
                        continue;
                    }
                    int timeTypeID = 1;
                    if (name.StartsWith("08:30")) 
                    {
                        timeTypeID = 1;
                    }
                    else if (name.StartsWith("12:30"))
                    {
                        timeTypeID = 2;
                    }
                    else if (name.StartsWith("19:30"))
                    {
                        timeTypeID = 3;
                    }

                    string tvRate = row[3].ToString();
                    decimal _tvRate = 0;
                    decimal.TryParse(tvRate, out _tvRate);

                    EachDayTVSeriesDO tvDO = new EachDayTVSeriesDO();
                    tvDO.StrTime = name;
                    tvDO.TimeTypeID = timeTypeID;
                    tvDO.EffectiveDate = reportDate;
                    tvDO.TVRate = _tvRate;
                    tvDO.EffectiveDate = reportDate;
                    BusinessLogicBase.Default.Insert(tvDO);
                }

                MessageBox.Show("读取成功！");
            }
            catch (Exception ex)
            {

            }
        }

        private void ReadChannelRankFile(DateTime reportDate, string filePath) 
        {
            try
            {
                DataTable dtChannel = NPOIHelper.ImportExceltoDt(filePath, "时期", 1);
                if (dtChannel == null || dtChannel.Rows.Count <= 0)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                for (int i = 1; i < dtChannel.Rows.Count; i++)
                {
                    DataRow row = dtChannel.Rows[i];

                    string name = row[0].ToString().Trim();
                    string tvPercentage = row[1].ToString();
                    decimal _tvPercentage = 0;
                    decimal.TryParse(tvPercentage, out _tvPercentage);

                    ChannelPercentageRankDO rankDO = new ChannelPercentageRankDO();
                    rankDO.ChannelName = name;
                    rankDO.EffectiveDate = reportDate;
                    rankDO.TVPercentage = _tvPercentage;
                    BusinessLogicBase.Default.Insert(rankDO);
                }

                MessageBox.Show("读取成功！");
            }
            catch (Exception ex)
            {

            }
        }

        private void ReadSpecialProgramFile(DateTime reportDate, string filePath) 
        {
            try
            {
                DataTable dtProgram = NPOIHelper.ImportExceltoDt(filePath, "交互分析", 1);
                if (dtProgram == null || dtProgram.Rows.Count <= 0)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                string programName = Path.GetFileNameWithoutExtension(filePath);
                ReadSpecialProgramSheet(programName, reportDate, dtProgram);

                MessageBox.Show("读取成功！");
            }
            catch (Exception ex)
            {

            }
        }

        private void ReadSpecialProgramSheet(string programName, DateTime reportDate, DataTable dtProgram)
        {
            DataRow row = dtProgram.Rows[2];
            EachProgramRatingDO eachProgramDO = DRManager.GetProgramRatingDOByName(programName, reportDate);
            if (eachProgramDO.ID > 0)
            {
                string tvRate = row[2].ToString();
                string tvPercentage = row[3].ToString();
                string tvLoyalty = row[4].ToString();
                decimal _tvRate = 0;
                decimal _tvPercentage = 0;
                decimal _tvLoyalty = 0;
                decimal.TryParse(tvRate, out _tvRate);
                decimal.TryParse(tvPercentage, out _tvPercentage);
                decimal.TryParse(tvLoyalty, out _tvLoyalty);

                eachProgramDO.TVRate = _tvRate;
                eachProgramDO.TVPercentage = _tvPercentage;
                eachProgramDO.TVLoyalty = _tvLoyalty;
                BusinessLogicBase.Default.Update(eachProgramDO);
            }
        }

        private void ReadEachMinuteProgramFile(DateTime reportDate, string filePath) 
        {
            try
            {
                DataTable dtEachMinute = NPOIHelper.ImportExceltoDt(filePath, "每分钟节目", 1);
                if (dtEachMinute == null || dtEachMinute.Rows.Count <= 0)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                DataTable dtEachProgram = NPOIHelper.ImportExceltoDt(filePath, "交互分析", 1);
                if (dtEachProgram == null || dtEachProgram.Rows.Count <= 0)
                {
                    MessageBox.Show("文件中未找到相关数据");
                    return;
                }
                ReadEachMinuteSheet(reportDate, dtEachMinute);
                ReadEachProgramSheet(reportDate, dtEachProgram);

                MessageBox.Show("读取成功！");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void ReadEachMinuteSheet(DateTime reportDate, DataTable dtEachMinute)
        {
            int currentMinuteIndex = 0;
            string currentProgramName = "";
            EachProgramRatingDO programDO = new EachProgramRatingDO();
            for (int i = 1; i < dtEachMinute.Rows.Count; i++)
            {
                DataRow row = dtEachMinute.Rows[i];
                string name = row[0].ToString().Trim();
                string tvRate = row[5].ToString();
                string tvPercentage = row[6].ToString();
                string tvLoyalty = row[7].ToString();
                decimal _tvRate = 0;
                decimal _tvPercentage = 0;
                decimal _tvLoyalty = 0;
                decimal.TryParse(tvRate, out _tvRate);
                decimal.TryParse(tvPercentage, out _tvPercentage);
                decimal.TryParse(tvLoyalty, out _tvLoyalty);
                if (name.Contains("<<") && name.Contains(">>"))
                {
                    currentMinuteIndex++;

                    string strMinute = name.Trim(new char[] { ' ', '<', '>' });
                    decimal rate = 0;
                    decimal.TryParse(tvRate, out rate);

                    EachMinuteRatingDO minuteDO = new EachMinuteRatingDO();
                    minuteDO.ProgramName = currentProgramName;
                    minuteDO.StrMinute = strMinute;
                    minuteDO.MinuteIndex = currentMinuteIndex;
                    minuteDO.EffectiveDate = reportDate;
                    minuteDO.TVRate = _tvRate;
                    minuteDO.TVPercentage = _tvPercentage;
                    minuteDO.TVLoyalty = _tvLoyalty;
                    BusinessLogicBase.Default.Insert(minuteDO);
                }
                else
                {
                    string strBeginTime = row[3].ToString().Trim(new char[] { ' ', '<', '>' });
                    string strEndTime = row[4].ToString().Trim(new char[] { ' ', '<', '>' });
                    int beginSeconds = CommonFunction.GetTimeToSeconds(strBeginTime);
                    int endSeconds = CommonFunction.GetTimeToSeconds(strEndTime);
                    if (beginSeconds > 24 * 3600) 
                    {
                        //only calculate program started before 23:00
                        break;
                    }
                    string matchedProgramName = GetMatchedSpecialProgram(name);
                    if (!string.IsNullOrEmpty(matchedProgramName))
                    {
                        //need to handle special program case
                        name = matchedProgramName;
                    }
                    programDO = DRManager.GetProgramRatingDOByName(name, reportDate, beginSeconds);
                    if (programDO != null && string.IsNullOrEmpty(programDO.ProgramName) == false)
                    {
                        programDO.StrEndTime = strEndTime;
                        programDO.EndSecond = endSeconds;
                        BusinessLogicBase.Default.Update(programDO);
                    }
                    else
                    {
                        programDO = new EachProgramRatingDO();
                        programDO.EffectiveDate = reportDate;
                        programDO.ProgramName = name;
                        programDO.StrBeginTime = strBeginTime;
                        programDO.StrEndTime = strEndTime;
                        programDO.BeginSecond = beginSeconds;
                        programDO.EndSecond = endSeconds;
                        programDO.TVRate = _tvRate;
                        programDO.TVPercentage = _tvPercentage;
                        programDO.TVLoyalty = _tvLoyalty;
                        BusinessLogicBase.Default.Insert(programDO);

                        currentMinuteIndex = 0;
                    }
                    currentProgramName = name;
                }
            }
        }

        private void ReadEachProgramSheet(DateTime reportDate, DataTable dtEachProgram) 
        {
            for (int i = 2; i < dtEachProgram.Rows.Count; i++)
            {
                DataRow row = dtEachProgram.Rows[i];
                string name = row[1].ToString().Trim();
                string channelName = row[2].ToString();
                string categoryName = row[3].ToString().Trim();
                string subCategoryName = row[4].ToString().Trim();

                string matchedSpecialProgram = GetMatchedSpecialProgram(name);
                if (!string.IsNullOrEmpty(matchedSpecialProgram)) 
                {
                    EachProgramRatingDO programDO = DRManager.GetProgramRatingDOByName(matchedSpecialProgram, reportDate);
                    if (programDO != null && programDO.ID > 0)
                    {
                        programDO.ChannelName = channelName;
                        programDO.CategoryName = categoryName;
                        programDO.SubCategoryName = subCategoryName;
                        BusinessLogicBase.Default.Update(programDO);
                    }
                    continue;
                }
                DataTable dt = DRManager.GetProgramRatingsByName(name, reportDate);
                if (dt != null && dt.Rows.Count == 1)
                {
                    int id = 0;
                    int.TryParse(dt.Rows[0]["ID"].ToString(), out id);
                    EachProgramRatingDO eachProgramDO = DRManager.GetProgramRatingDOByID(id);
                    if (eachProgramDO != null && eachProgramDO.ID > 0)
                    {
                        string tvRate = row[5].ToString();
                        string tvPercentage = row[6].ToString();
                        string tvLoyalty = row[7].ToString();
                        decimal _tvRate = 0;
                        decimal _tvPercentage = 0;
                        decimal _tvLoyalty = 0;
                        decimal.TryParse(tvRate, out _tvRate);
                        decimal.TryParse(tvPercentage, out _tvPercentage);
                        decimal.TryParse(tvLoyalty, out _tvLoyalty);
                        eachProgramDO.ChannelName = channelName;
                        eachProgramDO.CategoryName = categoryName;
                        eachProgramDO.SubCategoryName = subCategoryName;
                        eachProgramDO.TVRate = _tvRate;
                        eachProgramDO.TVPercentage = _tvPercentage;
                        eachProgramDO.TVLoyalty = _tvLoyalty;
                        BusinessLogicBase.Default.Update(eachProgramDO);
                    }
                }
                else if (dt != null && dt.Rows.Count > 1) 
                {
                    foreach (DataRow rateRow in dt.Rows) 
                    {
                        int id = 0;
                        int.TryParse(rateRow["ID"].ToString(), out id);
                        EachProgramRatingDO eachProgramDO = DRManager.GetProgramRatingDOByID(id);
                        if (eachProgramDO != null && eachProgramDO.ID > 0)
                        {
                            eachProgramDO.ChannelName = channelName;
                            eachProgramDO.CategoryName = categoryName;
                            eachProgramDO.SubCategoryName = subCategoryName;
                            BusinessLogicBase.Default.Update(eachProgramDO);
                        }
                    }
                } 
                
            }
        }

        private List<string> GetSystemSpecialPrograms() 
        {
            string programs = ConfigurationManager.AppSettings["SpecialPrograms"];
            string[] programList = programs.Split('|');
            return programList.ToList<string>();
        }

        private string GetMatchedSpecialProgram(string program) 
        {
            string matchedSPName = string.Empty;
            string programs = ConfigurationManager.AppSettings["SpecialPrograms"];
            string[] programList = programs.Split('|');
            foreach (string spName in programList) 
            {
                if (program.Contains(spName)) 
                {
                    matchedSPName = spName;
                    break;
                }
            }
            return matchedSPName;
        }

        private void btnCreateExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfdSaveFile = new SaveFileDialog();
            sfdSaveFile.Title = "保存文件";
            sfdSaveFile.Filter = "xlsx文件(*.xlsx)|*.xlsx";
            sfdSaveFile.RestoreDirectory = false;
            if (sfdSaveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string savePath = sfdSaveFile.FileName;
                DateTime reportDate = calReportDate.Value;
                DataTable dt = DRManager.GetAllMinutesRate(reportDate);

                DataTable dtOutput = new DataTable();
                dtOutput.TableName = "节目";
                dtOutput.Columns.Add("名称", typeof(string));
                dtOutput.Columns.Add("频道", typeof(string));
                dtOutput.Columns.Add("日期", typeof(string));
                dtOutput.Columns.Add("周日", typeof(string));
                dtOutput.Columns.Add("开始时间", typeof(string));
                dtOutput.Columns.Add("时长", typeof(string));
                dtOutput.Columns.Add("结束时间", typeof(string));
                dtOutput.Columns.Add("收视率%", typeof(string));
                dtOutput.Columns.Add("市场份额%", typeof(string));
                dtOutput.Columns.Add("平均忠实度", typeof(string));
                int columnNum = dt.Columns.Count - 12;
                for (int i = 1; i <= columnNum; i++)
                {
                    string minuteColumn = i + "分钟";
                    dtOutput.Columns.Add(minuteColumn, typeof(string));
                }

                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        DataRow newRow = dtOutput.NewRow();
                        newRow["名称"] = row["ProgramName"].ToString();
                        newRow["频道"] = row["ChannelName"].ToString();

                        DateTime effctiveDate = Constants.Date_Min;
                        DateTime.TryParse(row["EffectiveDate"].ToString(), out effctiveDate);
                        newRow["日期"] = effctiveDate.ToString("yyyy/M/d");

                        newRow["周日"] = row["WeekDayName"].ToString().Replace("星期", "周");
                        newRow["开始时间"] = row["StrBeginTime"].ToString();
                        newRow["时长"] = row["StrTimeLength"].ToString();
                        newRow["结束时间"] = row["StrEndTime"].ToString();
                        newRow["收视率%"] = row["TVRate"].ToString();
                        newRow["市场份额%"] = row["TVPercentage"].ToString();
                        newRow["平均忠实度"] = row["TVLoyalty"].ToString();

                        for (int i = 1; i <= columnNum; i++)
                        {
                            string minuteColumn = i + "分钟";
                            newRow[minuteColumn] = row[minuteColumn].ToString();
                        }

                        dtOutput.Rows.Add(newRow);
                    }

                    //foreach (DataColumn dc in dtOutput.Columns)
                    //{
                    //    dc.ColumnName = dc.ColumnName + ParseExcel.CellStyle.A_CA;
                    //}
                }

                NPOIHelper.ExportDTtoExcel(dtOutput, "", savePath);
                MessageBox.Show("输出成功！");
            }
        }

        private void btnCreateWord_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfdSaveFile = new SaveFileDialog();
            sfdSaveFile.Title = "保存文件";
            sfdSaveFile.Filter = "word文件(*.docx)|*.docx";
            sfdSaveFile.RestoreDirectory = false;
            if (sfdSaveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string outputPath = sfdSaveFile.FileName;
                DateTime reportDate = calReportDate.Value;
                string strReportDate = reportDate.ToString("yyyy年MM月dd日");
                string templatePath = Environment.CurrentDirectory + "\\Template.docx";
                File.Copy(templatePath, outputPath, true);

                //获取基础数据表
                DataTable dtRatingShare = DRManager.GetChannelPercentageRank(reportDate);
                DataTable dtRatingTarget = DRManager.GetProgramTargetRateRank(reportDate);
                //第一部分描述
                string cctv1rateDesc = GetCCTVPercentageDesc(dtRatingShare);
                string targetRateDesc = GetCCTVProgramDesc(dtRatingTarget, reportDate);
                //第二部分描述
                Dictionary<string, string> bookMappings = new Dictionary<string, string>();
                bookMappings.Add("ReportDate", strReportDate);
                bookMappings.Add("CCTV1RateDesc", cctv1rateDesc);
                bookMappings.Add("TargetRateDesc", targetRateDesc);
                //第四部分描述
                string weekendDesc = "";
                if (reportDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    weekendDesc = GetWeekendProgramDesc(reportDate);
                }
                bookMappings.Add("WeekendProgramRate", weekendDesc);
                DailyReport.WordCalc.CreateParagraph(outputPath, bookMappings);
                //第一部分表格
                DataTable dtChannelRank = dtRatingShare.Copy();
                dtChannelRank.Columns.Add("changeRateDesc", typeof(string));
                dtChannelRank.Columns.Add("changeRankDesc", typeof(string));
                foreach (DataRow row in dtChannelRank.Rows)
                {
                    int changeRank = 0;
                    decimal changeRate = 0;
                    int.TryParse(row["changeRank"].ToString(), out changeRank);
                    decimal.TryParse(row["changeRate"].ToString(), out changeRate);
                    int changeRate_int = Convert.ToInt32(changeRate);
                    row["changeRateDesc"] = changeRate_int.ToString() + "%";
                    if (changeRank > 0)
                    {
                        row["changeRankDesc"] = "↓" + changeRank;
                    }
                    else if (changeRank < 0)
                    {
                        row["changeRankDesc"] = "↑" + (-changeRank).ToString();
                    }
                    else
                    {
                        row["changeRankDesc"] = "0";
                    }
                }
                dtChannelRank.Columns.Remove("changeRank");
                dtChannelRank.Columns.Remove("changeRate");

                float[] widths = { 45f, 90f, 80f, 100f, 100f };
                string[] columns = { "排名", "频道", "市场份额%", "较前一日变化幅度", "较前一日排名变化" };
                string tableMark = "RatingShareTable";
                DailyReport.WordCalc.CreateTable(tableMark, outputPath, dtChannelRank, columns, widths);

                //第二部分表格
                DataTable dtTargetRank = dtRatingTarget.Copy();
                dtTargetRank.Columns.Remove("ChangeRate");
                dtTargetRank.Columns.Remove("BeginSecond");
                dtTargetRank.Columns.Remove("CategoryName");
                dtTargetRank.Columns.Add("blank1").SetOrdinal(1);
                dtTargetRank.Columns.Add("blank2").SetOrdinal(3);
                dtTargetRank.Columns.Add("blank3").SetOrdinal(5);
                dtTargetRank.Columns.Add("blank4").SetOrdinal(7);
                float[] targetWidths = { 60f, 11f, 90f, 11f, 50f, 11f, 110f, 11f, 60f };
                string[] targetColumns = { "播出时间", "", "名称", "", "收视率%", "", "较前一日变化幅度率", "", "目标完成率" };
                string targetTableMark = "RatingTargetTable";
                DailyReport.WordCalc.CreateTable(targetTableMark, outputPath, dtTargetRank, targetColumns, targetWidths);
                MessageBox.Show("输出成功！");
            }
        }

        private string GetCCTVProgramDesc(DataTable dtRank, DateTime reportDate)
        {
            string desc = reportDate.Day +"日";
            //上升的节目
            DataRow[] upRows = dtRank.Select("ChangeRate>0", "ChangeRate desc");
            if (upRows != null && upRows.Length > 0)
            {
                string upDesc = "栏目较前一日收视率提升的有：";
                string changeDesc  = GetChangedPrograms(upRows);
                if (string.IsNullOrEmpty(changeDesc) == false) 
                {
                    desc = desc + upDesc + changeDesc + "；";
                }
            }
            //持平的节目
            upRows = dtRank.Select("ChangeRate=0", "ChangeRate desc");
            if (upRows != null && upRows.Length > 0)
            {
                string upDesc = "栏目较前一日收视率持平的有：";
                string changeDesc = GetChangedPrograms(upRows);
                if (string.IsNullOrEmpty(changeDesc) == false)
                {
                    desc = desc + upDesc + changeDesc + "；";
                }
            }
            upRows = dtRank.Select("ChangeRate<0", "ChangeRate desc");
            if (upRows != null && upRows.Length > 0)
            {
                string upDesc = "栏目较前一日收视率下降的有：";
                string changeDesc = GetChangedPrograms(upRows);
                if (string.IsNullOrEmpty(changeDesc) == false)
                {
                    desc = desc + upDesc + changeDesc + "；";
                }
            }
            desc = desc.TrimEnd('；') + "。";
            return desc;
        }

        private static string GetChangedPrograms(DataRow[] changeRows)
        {
            char splitChar = '、';
            string changeDesc = "";
            foreach (DataRow row in changeRows)
            {
                string programName = row["ProgramName"].ToString();
                string categoryName = row["CategoryName"].ToString();
                if (categoryName == "电视剧")
                {
                    int beginSecond = row["BeginSecond"].ToString().ToInt(0);
                    if (beginSecond < 12 * 3600)
                    {
                        programName = "上午剧场《" + programName + "》";
                    }
                    else if (beginSecond < 19 * 3600)
                    {
                        programName = "下午剧场《" + programName + "》";
                    }
                    else if (beginSecond < 22 * 3600)
                    {
                        programName = "晚间黄金剧场《" + programName + "》";
                    }
                    else
                    {
                        programName = "《" + programName + "》";
                    }
                }
                else
                {
                    programName = "《" + programName + "》";
                }

                changeDesc = changeDesc + programName + splitChar;
            }
            changeDesc = changeDesc.TrimEnd(splitChar);
            if (changeRows.Length > 1)
            {
                int lastSplitCharPosition = changeDesc.LastIndexOf(splitChar);
                changeDesc = changeDesc.Remove(lastSplitCharPosition, 1);
                changeDesc = changeDesc.Insert(lastSplitCharPosition, "和");
            }

            return changeDesc;
        }

        private string GetWeekendProgramDesc(DateTime reportDate) 
        {
            StringBuilder descSb = new StringBuilder();
            DateTime fromDate = reportDate.AddDays(-1);
            DateTime toDate = reportDate;
            DataTable dt = DRManager.GetWeekendPrograms(fromDate, toDate);
            if (dt != null && dt.Rows.Count > 0) 
            {
                foreach (DataRow row in dt.Rows) 
                {
                    string programName = row["ProgramName"].ToString();
                    string effectiveDate = row["EffectiveDate"].ToString();
                    string strBeginTime = row["StrBeginTime"].ToString();
                    string strEndTime = row["StrEndTime"].ToString();
                    string tvRate = row["TVRate"].ToString();

                    DateTime _effectiveDate = Constants.Date_Min;
                    DateTime.TryParse(effectiveDate, out _effectiveDate);
                    string strDate = _effectiveDate.ToString("yyyy年M月d日") + strBeginTime.Substring(0, 5) + "-" + strEndTime.Substring(0, 5);

                    
                    decimal _tvRate = 0;
                    decimal.TryParse(tvRate, out _tvRate);
                    decimal lastRate = DRManager.GetWeekendProgramsByName(_effectiveDate.AddDays(-7), programName);
                    string gapDesc = "";
                    if (lastRate > 0)
                    {
                        decimal gapValue = 100 * (_tvRate - lastRate) / lastRate;
                        if (gapValue > 0)
                        {
                            gapDesc = string.Format("上升{0}%", Math.Abs(gapValue).ToString("N1"));
                        }
                        else if (gapValue < 0)
                        {
                            gapDesc = string.Format("下降{0}%", Math.Abs(gapValue).ToString("N1"));
                        }
                        else
                        {
                            gapDesc = "持平";
                        }
                    }

                    string desc = string.Format("{0}《{1}》【34城市】收视率为（{2}%）{3}。", strDate, programName, _tvRate.ToString("N4"), gapDesc);
                    descSb.AppendLine(desc);
                }
            }
            return descSb.ToString();
        }

        private string GetCCTVPercentageDesc(DataTable dtRatingShare) 
        {
            DataRow row = dtRatingShare.Select("ChannelName='综合频道'")[0];
            string rank = row["RowNum"].ToString();
            string tvPercentage = row["TVPercentage"].ToString();
            string changeRate = row["changeRate"].ToString();
            string changeRank = row["changeRank"].ToString();
            decimal gap = 0;
            decimal.TryParse(changeRate, out gap);
            string gapDesc = Math.Abs(gap).ToString("N0");
            //综合频道收视份额2.9933%，较前一日下降4.7%，排名全国第六位，与前一日排位相同。
            string desc = string.Format("综合频道收视份额{0}%，", tvPercentage);
            if (gap > 0)
            {
                desc = desc + string.Format("较前一日上升{0}%，", gapDesc);
            }
            else if (gap < 0)
            {
                desc = desc + string.Format("较前一日下降{0}%，", gapDesc);
            }
            else 
            {
                desc = desc + "与前一日持平，";
            }

            int rankGap = 0;
            int.TryParse(changeRank, out rankGap);
            int rankGapDesc = Math.Abs(rankGap);
            desc = desc + string.Format("排名全国第{0}位，", rank);
            if (rankGap < 0)
            {
                desc = desc + string.Format("较前一日排名上升{0}位。", rankGapDesc);
            }
            else if (rankGap > 0)
            {
                desc = desc + string.Format("较前一日排名下降{0}位。", rankGapDesc);
            }
            else
            {
                desc = desc + "与前一日排名相同。";
            }
            /*
            //上午8:38 《生活圈》(0.2042%)下降1.1%。9:27《飞虎队》（0.3203%）下降8.5%；
            DataTable morningPrograms = DRManager.GetMorningPrograms(reportDate);
            DataTable afternoonPrograms = DRManager.GetAfternoonPrograms(reportDate);
            DataTable evenningPrograms = DRManager.GetEvenningPrograms(reportDate);
            desc += "上午";
            CompareEachProgram(reportDate, ref desc, morningPrograms,1);
            desc += "下午";
            CompareEachProgram(reportDate, ref desc, afternoonPrograms,2);
            CompareEachProgram(reportDate, ref desc, evenningPrograms,3);
            desc = desc.Trim().TrimEnd('；') + "。";
            */
            return desc;
        }

        private static void CompareEachProgram(DateTime reportDate, ref string desc, DataTable morningPrograms, int timeTypeID)
        {
            if (morningPrograms != null && morningPrograms.Rows.Count > 0)
            {
                bool bHasHandleSeries = false;
                foreach (DataRow pRow in morningPrograms.Rows)
                {
                    string programName = pRow["ProgramName"].ToString();
                    string categoryName = pRow["CategoryName"].ToString();
                    string strBeginTime = pRow["StrBeginTime"].ToString();
                    string beginSecond = pRow["BeginSecond"].ToString();
                    string endSecond = pRow["EndSecond"].ToString();
                    string thisRate = pRow["TVRate"].ToString();
                    int _beginSecond = 0;
                    int _endSecond = 0;
                    int.TryParse(beginSecond, out _beginSecond);
                    int.TryParse(endSecond, out _endSecond);

                    decimal tvRate = 0;
                    decimal lastRate = 0;

                    //if (categoryName == "电视剧")
                    //{
                    //    if (bHasHandleSeries == true)
                    //    {
                    //        continue;
                    //    }
                    //    //int timeTypeID = 1;
                    //    //if (_beginSecond <= (11 * 3600 + 1800))
                    //    //{
                    //    //    //8:30-11:30时段的电视剧
                    //    //    timeTypeID = 1;
                    //    //}
                    //    //else if (_beginSecond >= (12 * 3600 + 1800) && _beginSecond <= (18 * 3600 + 1800))
                    //    //{
                    //    //    //12:30-18:00时段的电视剧
                    //    //    timeTypeID = 2;
                    //    //}
                    //    //else if (_beginSecond >= (19 * 3600 + 1800))
                    //    //{
                    //    //    //19:30-22:00时段的电视剧
                    //    //    timeTypeID = 3;
                    //    //}

                    //    EachDayTVSeriesDO tvDO = DRManager.GetTVSeries(reportDate, timeTypeID);
                    //    if (tvDO != null && tvDO.ID > 0)
                    //    {
                    //        tvRate = tvDO.TVRate;
                    //        lastRate = DRManager.GetLastSameTVSeriesRate(reportDate, timeTypeID);
                    //    }
                    //    else
                    //    {
                    //        continue;
                    //    }
                    //    bHasHandleSeries = true;
                    //}
                    //else
                    if (programName == "第1动画乐园" || programName == "生活圈")
                    {
                        decimal.TryParse(thisRate, out tvRate);
                        lastRate = DRManager.GetLastSameProgramRate(reportDate, programName, _beginSecond);
                    }
                    else
                    {
                        decimal.TryParse(thisRate, out tvRate);
                        lastRate = DRManager.GetLastSameProgramRateInSameDay(reportDate, programName, _beginSecond);
                    }

                    decimal gapValue = 0;
                    string gapDescPart2 = "";
                    if (lastRate > 0)
                    {
                        gapValue = 100 * (tvRate - lastRate) / lastRate;
                        gapDescPart2 = Math.Abs(gapValue).ToString("N1");
                    }
                    string strTVRate = tvRate.ToString("N4");
                    if (gapValue > 0)
                    {
                        desc = desc + string.Format("{0}《{1}》({2}%)上升{3}%； ", strBeginTime.Substring(0, 5), programName, strTVRate, gapDescPart2);
                    }
                    else if (gapValue < 0)
                    {
                        desc = desc + string.Format("{0}《{1}》({2}%)下降{3}%； ", strBeginTime.Substring(0, 5), programName, strTVRate, gapDescPart2);
                    }
                    else
                    {
                        gapDescPart2 = "持平";
                        if (lastRate == 0)
                        {
                            gapDescPart2 = "";
                        }
                        desc = desc + string.Format("{0}《{1}》({2}%){3}； ", strBeginTime.Substring(0, 5), programName, strTVRate, gapDescPart2);
                    }

                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            DateTime reportDate = calReportDate.Value;
            DialogResult dr = MessageBox.Show(string.Format("您确认要删除{0}的所有元数据吗？", reportDate.ToString("yyyy年MM月dd日")), "确认", MessageBoxButtons.YesNo);
            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                DRManager.ClearDataByDate(reportDate);
                MessageBox.Show("清除成功！");
            }
        }

        private void LoadData()
        {
            DateTime reportDate = calReportDate.Value;
            if (ddlFileType.SelectedItem == null) 
            {
                this.gvData.DataSource = null;
                this.gvData.Update();
                return;
            }
            string fileTypeName = ddlFileType.SelectedItem.ToString();
            if (string.IsNullOrEmpty(fileTypeName)) 
            {
                this.gvData.DataSource = null;
                this.gvData.Update();
                return;
            }
            this.gvData.DataSource = DRManager.GetUploadedDataByDate(reportDate,fileTypeName);
            this.gvData.Update();
        }

        private void ddlFileType_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadData();
        }

        private void calReportDate_ValueChanged(object sender, EventArgs e)
        {
            LoadData();
        }

    }
}
