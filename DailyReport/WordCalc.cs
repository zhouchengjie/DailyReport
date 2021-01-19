using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Data;

namespace DailyReport
{
    public static class WordCalc
    {
        public static void CreateDoc(string filePath, string outputPath, Dictionary<string, string> bookmappings)
        {
            System.IO.File.Copy(filePath, outputPath, true);

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = outputPath;//最终的word文档需要写入的位置  
            object ModelName = outputPath;//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  


            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板  

            try
            {
                foreach (KeyValuePair<string, string> pair in bookmappings)
                {
                    string bookmark = pair.Key;
                    string value = pair.Value;
                    object oStart = bookmark;//word中的书签名   
                    try
                    {
                        Microsoft.Office.Interop.Word.Range range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//表格插入位置   
                        range.Text = value;//在书签处插入文字内容  
                    }
                    catch (Exception ex)
                    {

                    }
                }

                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;//保存格式   
                wordDoc.SaveAs(ref templateName, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch
            {

            }
            finally
            {
                //关闭wordDoc，wordApp对象                
                object SaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object OriginalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                app.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
        }

        public static void CreateGraph()
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = @"C:\Users\zhouchengjie\Desktop\2017-07-17\test.docx";//最终的word文档需要写入的位置  
            object ModelName = @"C:\Users\zhouchengjie\Desktop\2017-07-17\test.docx";//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  
            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板  

            //UpdateGraph1(ref wordDoc, lastValues, currentValues);

            app.Quit(true);

        }

        public static void UpdateGraph1(string outputPath, double[] lastValues, double[] currentValues, int graphIndex)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = outputPath;//最终的word文档需要写入的位置  
            object ModelName = outputPath;//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  


            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板  

            string[] titles = { "日期", "本周", "上周" };
            string[] names = { "周五", "周六", "周日", "周一", "周二", "周三", "周四" }; // 数据名称 
            int fieldCount = names.Length;

            try
            {
                Microsoft.Office.Interop.Word.InlineShape shape = wordDoc.InlineShapes[graphIndex];
                Microsoft.Office.Interop.Word.Chart chart = shape.Chart;
                chart.ChartData.Activate();
                Microsoft.Office.Interop.Excel.Worksheet book = chart.ChartData.Workbook.Worksheets["Sheet1"];

                var data = new object[fieldCount, 3];
                for (int i = 0; i < fieldCount; i++)
                {
                    data[i, 0] = names[i];
                    data[i, 1] = currentValues[i];
                    data[i, 2] = lastValues[i];
                }
                book.get_Range("A2", "C" + (fieldCount + 1)).Value = data;
                book.get_Range("A1", "C1").Value = titles;
                ((Microsoft.Office.Interop.Excel.Range)book.Cells[1, titles.Length + 1]).Select();
                ((Microsoft.Office.Interop.Excel.Range)book.Cells[1, titles.Length + 1]).EntireColumn.Delete(0);

                book.Application.DisplayAlerts = false;
                book.Application.Quit();

                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;//保存格式   
                wordDoc.SaveAs(ref templateName, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch
            {

            }
            finally
            {
                //关闭wordDoc，wordApp对象                
                object SaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object OriginalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                app.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
        }

        public static void UpdateGraph2(string outputPath, string[] names, double[] currentValues, int graphIndex)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = outputPath;//最终的word文档需要写入的位置  
            object ModelName = outputPath;//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  


            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板


            try
            {
                int fieldCount = names.Length;
                Microsoft.Office.Interop.Word.InlineShape shape = wordDoc.InlineShapes[graphIndex];
                Microsoft.Office.Interop.Word.OLEFormat oleformat = shape.OLEFormat;
                oleformat.Open();
                Microsoft.Office.Interop.Graph.Chart chart = (Microsoft.Office.Interop.Graph.Chart)(oleformat.Object);
                chart.Application.DataSheet.Cells.ClearContents();
                chart.Application.DataSheet.Cells.set_Item(1, "", "名称");
                for (int i = 0; i < names.Length; i++)
                {
                    chart.Application.DataSheet.Cells.set_Item(i + 2, "", names[i]);
                    chart.Application.DataSheet.Cells.set_Item(i + 2, "A", currentValues[i]);
                }
                chart.Application.Update();
                chart.Application.DisplayAlerts = false;
                chart.Application.Quit();

                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;//保存格式   
                wordDoc.SaveAs(ref templateName, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch
            {

            }
            finally
            {
                //关闭wordDoc，wordApp对象                
                object SaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object OriginalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                app.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
        }

        public static void CreateTable(string positionMark, string outputPath, System.Data.DataTable dt, string[] columns, float[] widths)
        {
            if(dt == null)
            {
                return;
            }

            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = outputPath;//最终的word文档需要写入的位置  
            object ModelName = outputPath;//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  


            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板  
            try
            {
                wordDoc.ActiveWindow.Visible = true;
                object oStart = positionMark;
                Bookmark bk = wordDoc.Bookmarks.get_Item(ref oStart);//表格插入位置  
                if (bk == null) 
                {
                    return;
                }
                int rowCount = dt.Rows.Count + 1;
                int columnCount = columns.Length;

                Range range = bk.Range;
                range.Tables.Add(range, rowCount, columnCount);
                Table tb = range.Tables[1];
                tb.LeftPadding = 0f;
                tb.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                tb.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;
                tb.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleSingle;

                for (int i = 0; i < columnCount; i++)
                {
                    tb.Columns[i + 1].Width = widths[i];
                    tb.Cell(1, i + 1).Range.Text = columns[i];
                }
                
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    DataRow row = dt.Rows[i];
                    for (int j = 0; j < columnCount; j++)
                    {
                        tb.Cell(i + 2, j + 1).Range.Text = row[j].ToString();
                        if (positionMark == "RatingTargetTable" && dt.Columns[j].ColumnName== "StrChangeRate") 
                        {
                            string changeRateDesc = row["StrChangeRate"].ToString();
                            if (!changeRateDesc.Contains("-"))
                            {
                                tb.Cell(i + 2, j + 1).Range.Font.Color = WdColor.wdColorRed;
                                tb.Cell(i + 2, j + 1).Range.Font.Bold = 1;
                            }
                        }
                    }
                    if (positionMark == "RatingShareTable") 
                    {
                        string channelName = row["ChannelName"].ToString();
                        if (channelName == "综合频道") 
                        {
                            tb.Rows[i + 2].Range.Shading.BackgroundPatternColor = WdColor.wdColorYellow;
                            tb.Rows[i + 2].Range.Font.Color = WdColor.wdColorRed;
                            tb.Rows[i + 2].Range.Font.Bold = 1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {

            }
            finally
            {
                //关闭wordDoc，wordApp对象                
                object SaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object OriginalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                app.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
        }

        public static void CreateParagraph(string outputPath, Dictionary<string, string> bookmappings)
        {
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            object missing = System.Reflection.Missing.Value;
            object templateName = outputPath;//最终的word文档需要写入的位置  
            object ModelName = outputPath;//word  模板的位置
            object count = 1;
            object WdLine = Microsoft.Office.Interop.Word.WdUnits.wdLine;//换一行;  


            Microsoft.Office.Interop.Word.Document wordDoc = app.Documents.Open(ref ModelName, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
               ref missing);//打开word模板  

            try
            {
                foreach (KeyValuePair<string, string> pair in bookmappings)
                {
                    string bookmark = pair.Key;
                    string value = pair.Value;
                    object oStart = bookmark;//word中的书签名   
                    try
                    {
                        Microsoft.Office.Interop.Word.Range range = wordDoc.Bookmarks.get_Item(ref oStart).Range;//表格插入位置   
                        range.Text = value;//在书签处插入文字内容  
                    }
                    catch (Exception ex)
                    {

                    }
                }

                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault;//保存格式   
                wordDoc.SaveAs(ref templateName, ref format, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            }
            catch
            {

            }
            finally
            {
                //关闭wordDoc，wordApp对象                
                object SaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object OriginalFormat = Microsoft.Office.Interop.Word.WdOriginalFormat.wdWordDocument;
                object RouteDocument = false;
                wordDoc.Close(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
                app.Quit(ref SaveChanges, ref OriginalFormat, ref RouteDocument);
            }
        }
    }
}
