using Microsoft.Extensions.Configuration;
using PRB.Repository.DataContext;
using PRB.Services;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.EntityFrameworkCore;
using Dapper;
using Microsoft.Data.SqlClient;
using Microsoft.Office.Interop.Word;
using Serilog;
using System.Globalization;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using System.Data;
using PRB.Repository.Repository;
using System.Reflection.Metadata.Ecma335;
using System.Security.Policy;
using System.Collections;
using System.IO;
using Microsoft.Office.Interop.Excel;
using Nest;
using System.Reflection;
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Text.RegularExpressions;
using HtmlAgilityPack;
using System.Web;

namespace PRB.Repository.Automation_Repository
{
    public interface IAutomation_Repository
    {
        public IDictionary<string, string> FetchUserName(string mailId);
        public string DetaialedReportGenerator(Word.Application oWord, Document oWordDoc, string monthYr, string file, int templateMatcher);
        public string SummaryReportGenerator(Document oWordDoc, string monthYr);
        public string DisclaimerReportGenerator(Document oWordDoc, string monthYr);
        public string MergeDoc(Word.Application oWord, string monthYr, string password);
        public string Generate(string report);
        public string SummaryFetcher(Word.Document oWordDoc, string monthYr);
        public string DisclaimerFetcher(Word.Document oWordDoc, string monthYr);
        public List<dynamic> StockDetails(int month, int year, string ticker);

    }
    public class Automation_Repository : IAutomation_Repository
    {
        private readonly IConfiguration? _configuration;
        private readonly PRB_DB_Context? _context;
        private string connectionStr = string.Empty;
        private readonly AppSettings? appSettings;
        protected readonly SqlConnection connection;

        protected readonly IRuleExecutor rulesExecutor;
        public Automation_Repository(IConfiguration configuration, IRuleExecutor rExecutor, PRB_DB_Context context)
        {
            _configuration = configuration;
            _context = context;
            connectionStr = configuration.GetConnectionString("MyDBConnection");
            this.connection = new SqlConnection(connectionStr);
            appSettings = configuration.Get<AppSettings>();
            this.rulesExecutor = rExecutor;
        }
        //Automation
        public string DetaialedReportGenerator(Word.Application oWord, Word.Document oWordDoc, string monthYr, string file, int templateMatcher)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;

            List<string> tickerlists = new List<string>();

            if (file == "Generate") tickerlists = GettingTicker(monthYr);
            else tickerlists.Add(file);



            int counter = 0;
            Redo:

            foreach (string ticker in tickerlists)
            {
                string templateId = TemplateIdProvider(ticker);
                if (templateMatcher != 0) templateId = templateMatcher.ToString();
                Object? oTemplatePath = appSettings.TemplatePath?.DetailedReportTemplatePath + templateId + ".docx";
                oWordDoc = oWord.Documents.Add(ref oTemplatePath);
                Log.Verbose("Template opened.");
                object? templatePassword = appSettings.Passwords?.TemplatePassword;
                oWordDoc.Unprotect(ref templatePassword);

                string synopsis = TickerSynopsis(ticker);
                if (synopsis == "") Log.Warning("Company description is null here.");

                List<dynamic> stockTable = StockDetails(month, year, ticker);
                if (stockTable.Count == 0) Log.Warning("Company transaction is null here.");

                List<dynamic> openClose = OCBalance(ticker, monthYr);

                var stockDetail = (from SC in _context.PrbSectorCodes
                                   join Tt in _context.PrbTickers on SC.SectorCode equals Tt.SectorCode
                                   join HD in _context.PrbHoldingDetails on Tt.CompanyTicker equals HD.CompanyTicker
                                   join Tk in _context.PrbTickers on HD.CompanyTicker equals Tk.CompanyTicker
                                   join CP in _context.PrbCompanyPrices on HD.CompanyTicker equals CP.CompanyTicker
                                   join CC in _context.PrbCurrencyCodes on HD.CurrencyCode equals CC.CurrencyCode
                                   where Tt.CompanyTicker == ticker && CP.CompanyTicker == ticker && CP.ReportDate.Month == month && HD.TransactionDate.Month == month && HD.TransactionDate.Year == year
                                   select new
                                   {
                                       company = Tt.CompanyName,
                                       sector = SC.SectorName,
                                       current = CP.LastMarketPrice,
                                       date = CP.ReportDate.ToShortDateString(),
                                       holdings = (from H in _context.PrbHoldingDetails where H.TransactionTypeCode == "B" && H.CompanyTicker == ticker && H.TransactionDate.Month <= month && H.TransactionDate.Year <= year select H.Quantity).Sum() - (from H in _context.PrbHoldingDetails where H.TransactionTypeCode == "S" && H.CompanyTicker == ticker && H.TransactionDate.Month <= month && H.TransactionDate.Year <= year select H.Quantity).Sum(),
                                       total = ((from H in _context.PrbHoldingDetails where H.TransactionTypeCode == "B" && H.CompanyTicker == ticker && H.TransactionDate.Month <= month && H.TransactionDate.Year <= year select H.Quantity).Sum() - (from H in _context.PrbHoldingDetails where H.TransactionTypeCode == "S" && H.CompanyTicker == ticker && H.TransactionDate.Month <= month && H.TransactionDate.Year <= year select H.Quantity).Sum()) * CP.LastMarketPrice,
                                       currency = CC.CurrencyDesc,
                                       currencyCode = CC.CurrencyCode
                                   }).FirstOrDefault();
                stockTable.Add(stockDetail.current);



                Log.Information("Binding the data in Word template...");

                //Word.ContentControl contentPieChart = (oWordDoc.SelectContentControlsByTag("contentPieChart"))[1];

                Word.Table tbl = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark12].Range.Tables[1];
                CCBinder(oWordDoc, "monthYr", monthYr);
                CCBinder(oWordDoc, "contentCompanyTicker", ticker);
                CCBinder(oWordDoc, "contentAboutCompany", synopsis);
                CCBinder(oWordDoc, "contentCompanyName", stockDetail.company);
                CCBinder(oWordDoc, "contentSector", stockDetail.sector);
                CCBinder(oWordDoc, "contentTicker", ticker);
                CCBinder(oWordDoc, "contentPrice", stockDetail.current.ToString());
                CCBinder(oWordDoc, "contentCurrentDate", stockDetail.date);
                CCBinder(oWordDoc, "contentHoldings", stockDetail.holdings.ToString());
                CCBinder(oWordDoc, "contentTotal", stockDetail.total.ToString("0.00"));
                CCBinder(oWordDoc, "contentCurrency", stockDetail.currency);
                CCBinder(oWordDoc, "contentCurrencyCode", stockDetail.currencyCode);
                CCBinder(oWordDoc, "contentOpening", openClose[0].opening.ToString("0.00"));
                CCBinder(oWordDoc, "contentClosing", openClose[0].closing.ToString("0.00"));

                //imgrng.InlineShapes.AddPicture("D:\\PRB.Services\\PRB.Services\\wwwroot\\Images\\Downward.jpg", oMissing, oMissing, imgrng);

                Log.Information("Data binded in Content Control.");

                try
                {
                    switch (templateId)
                    {
                        case "1": DetailedTableBinderV1(oWordDoc, tbl, stockTable); break;
                        case "2": DetailedTableBinderV2(oWordDoc, tbl, stockTable, openClose); break;
                        case "3": DetailedTableBinderV3(oWordDoc, tbl, stockTable); break;
                    }
                    counter++;
                    if (counter == 1) goto Redo;
                }
                catch (Exception ex)
                {
                    Log.Error("Couldn't bind data in chart.", ex);
                    throw new Exception("Couldn't bind data in chart.", ex);
                }
                finally
                {
                    Log.Debug($"Saving the document under directory {appSettings.ReportPath?.Path + ticker + ".docx..."}");
                    oWordDoc.SaveAs2(appSettings.ReportPath?.Path + ticker + ".docx");
                    Log.Verbose("Exporting the document as PDF format...");
                    oWordDoc.ExportAsFixedFormat(appSettings.ReportPath?.Path + ticker + ".pdf", WdExportFormat.wdExportFormatPDF);
                    oWordDoc.Close();
                    connection.Close();
                }
            }
            return null;
        }
        protected void CCBinder(Word.Document oWordDoc, string ccTag, string ccValue)
        {
            try
            {
                Word.ContentControl cc = null;
                cc = (oWordDoc.SelectContentControlsByTag(ccTag))[1];
                cc.Range.Text = ccValue;
            }
            catch (Exception ex)
            {
                Log.Error($"Coludn't bind data. {ex.Message}.\n");
                throw new Exception($" Coludn't bind data. {ex.Message}.\n");
            }
        }

        public string DetailedTableBinderV1(Word.Document oWordDoc, Word.Table tbl, List<dynamic> stockTable)
        {
            Word.ContentControl contentLineChart = (oWordDoc.SelectContentControlsByTag("contentLineChart"))[1];

            Word.InlineShape shape = contentLineChart.Range.InlineShapes[1];
            Word.Chart chart = shape.Chart;
            Excel.Workbook? wb = chart.ChartData.Workbook;
            Excel.Worksheet? ws = wb.Worksheets["Sheet1"];
            try
            {
                //Deletes the rows after 1st company until no.of rows = 3 including Heading
                if (tbl.Rows.Count > 3) for (int i = 1; tbl.Rows.Count > 3; i++) tbl.Rows[2].Delete();
                //Dynamically adds row in Word Table
                for (int i = 1; i < stockTable.Count - 1; i++) tbl.Rows.Add();
                int rowCounter = 2;
                foreach (var row in stockTable)
                {
                    tbl.Cell(rowCounter, 1).Range.Text = row.Transaction_Date.ToString("dd.MM.yyyy");
                    tbl.Cell(rowCounter, 2).Range.Text = row.Quantity.ToString();
                    tbl.Cell(rowCounter, 3).Range.Text = row.Amount.ToString();
                    tbl.Cell(rowCounter, 4).Range.Text = row.Transaction_Desc;
                    tbl.Cell(rowCounter, 5).Range.Text = row.total_Amount.ToString("0.00");



                    //Dynamically adds row in WorkBook
                    if (rowCounter < stockTable.Count - 1) ws.Rows[rowCounter].Insert();



                    ws.Cells[rowCounter, "A"].Value = row.Transaction_Date.ToString("dd.MM.yyyy"); Thread.Sleep(25);
                    ws.Cells[rowCounter, "B"].Value = row.Amount.ToString(); Thread.Sleep(25);



                    rowCounter++; if (rowCounter == stockTable.Count + 1) break;
                }
                return "Done";
            }
            catch (Exception ex)
            {
                Log.Error("Couldn't bind data in chart.", ex);
                throw new Exception("Couldn't bind data in chart.", ex);
            }
            finally
            {
                wb.Close(true);

                if (shape != null) Marshal.ReleaseComObject(shape);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                ws = null; wb = null;
                connection.Close();
            }

        }
        public string DetailedTableBinderV2(Word.Document oWordDoc, Word.Table tbl, List<dynamic> stockTable, List<dynamic> openClose)
        {
            Word.ContentControl worthChange = (oWordDoc.SelectContentControlsByTag("worthChange"))[1]; worthChange.Range.Delete();
            if (openClose[0].closing > openClose[0].opening) worthChange.Range.InlineShapes.AddPicture(appSettings.Images.Upward);
            else if (openClose[0].closing < openClose[0].opening) worthChange.Range.InlineShapes.AddPicture(appSettings.Images.Downward);
            else worthChange.Range.Text = "Nil";

            Word.ContentControl contentLineChart = (oWordDoc.SelectContentControlsByTag("contentLineChart"))[1];

            Word.InlineShape shape = contentLineChart.Range.InlineShapes[1];
            Word.Chart chart = shape.Chart;
            Excel.Workbook? wb = chart.ChartData.Workbook;
            Excel.Worksheet? ws = wb.Worksheets["Sheet1"];
            try
            {
                //Deletes the rows after 1st company until no.of rows = 3 including Heading
                if (tbl.Rows.Count > 3) for (int i = 1; tbl.Rows.Count > 3; i++) tbl.Rows[2].Delete();
                //Dynamically adds row in Word Table
                for (int i = 1; i < stockTable.Count - 1; i++) tbl.Rows.Add();
                int rowCounter = 2;
                foreach (var row in stockTable)
                {
                    tbl.Cell(rowCounter, 1).Range.Text = row.Transaction_Date.ToString("dd.MM.yyyy");
                    tbl.Cell(rowCounter, 2).Range.Text = row.Quantity.ToString();
                    tbl.Cell(rowCounter, 3).Range.Text = row.Amount.ToString();
                    tbl.Cell(rowCounter, 4).Range.Text = row.Transaction_Desc;
                    tbl.Cell(rowCounter, 5).Range.Text = row.total_Amount.ToString("0.00");
                    tbl.Cell(rowCounter, 6).Range.Text = (row.Quantity * stockTable[^1]).ToString("0.00");

                    Word.Range rng = tbl.Cell(rowCounter, 7).Range;
                    if (((((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100) > 0)
                    {
                        rng.Delete();
                        rng.InlineShapes.AddPicture(appSettings.Images.UpArrow); Thread.Sleep(25);
                        tbl.Cell(rowCounter, 8).Range.Text = ((((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100).ToString("+ 00.00");
                    }
                    else if (((((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100) < 0)
                    {
                        rng.Delete();
                        rng.InlineShapes.AddPicture(appSettings.Images.DownArrow); Thread.Sleep(25);
                        tbl.Cell(rowCounter, 8).Range.Text = (Math.Abs(((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100).ToString("- 00.00");
                    }
                    else
                    {
                        rng.Delete();
                        tbl.Cell(rowCounter, 7).Range.Text = "Nil";
                        tbl.Cell(rowCounter, 8).Range.Text = (Math.Abs(((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100).ToString("  00.00");

                    }

                    //Word.Range rng = tbl.Cell(rowCounter, 7).Range;
                    //if (((((row.Quantity * stockTable[^1]) - row.total_Amount) / row.total_Amount) * 100)<0) 
                    //    rng.InlineShapes.AddPicture(@"C:\Users\vmariappan\Downloads\images.png");
                    //tbl.Cell(rowCounter, 7).Range.Text = (((stockTable[^1] - row.Amount) / row.Amount) * 100).ToString("0.00");



                    //Dynamically adds row in WorkBook
                    if (rowCounter < stockTable.Count - 1) ws.Rows[rowCounter].Insert();



                    ws.Cells[rowCounter, "A"].Value = row.Transaction_Date.ToString("dd.MM.yyyy"); Thread.Sleep(25);
                    ws.Cells[rowCounter, "B"].Value = row.total_Amount.ToString("0.00"); Thread.Sleep(25);
                    ws.Cells[rowCounter, "C"].Value = (row.Quantity * stockTable[^1]).ToString("0.00"); Thread.Sleep(25);



                    rowCounter++; if (rowCounter == stockTable.Count + 1) break;
                }
                return "Done";
            }
            catch (Exception ex)
            {
                Log.Error("Couldn't bind data in chart.", ex);
                throw new Exception("Couldn't bind data in chart.", ex);
            }
            finally
            {
                wb.Close(true);

                if (shape != null) Marshal.ReleaseComObject(shape);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                ws = null; wb = null;
                connection.Close();
            }

        }
        public string DetailedTableBinderV3(Word.Document oWordDoc, Word.Table tbl, List<dynamic> stockTable)
        {
            Word.ContentControl contentColumnChart = (oWordDoc.SelectContentControlsByTag("contentColumnChart"))[1];
            Word.ContentControl contentLineChart = (oWordDoc.SelectContentControlsByTag("contentLineChart"))[1];

            Word.InlineShape shape1 = contentColumnChart.Range.InlineShapes[1];
            Word.Chart chart1 = shape1.Chart;
            Excel.Workbook? wb1 = chart1.ChartData.Workbook;
            Excel.Worksheet? ws1 = wb1.Worksheets["Sheet1"];

            Word.InlineShape shape2 = contentLineChart.Range.InlineShapes[1];
            Word.Chart chart2 = shape2.Chart;
            Excel.Workbook? wb2 = chart2.ChartData.Workbook;
            Excel.Worksheet? ws2 = wb2.Worksheets["Sheet1"];
            List<List<string>> table = new List<List<string>>();
            int counter = 0;
            foreach (var row in stockTable)
            {
                table.Add(new List<string> { row.Transaction_Date.ToString("dd.MM.yyyy"), row.Quantity.ToString(), row.Amount.ToString(), row.Transaction_Desc, row.total_Amount.ToString("0.00") });
                counter++;
                if (counter == stockTable.Count - 1) break;
            }
            try
            {
                //Deletes the rows after 1st company until no.of rows = 3 including Heading
                if (tbl.Rows.Count > 3) for (int i = 1; tbl.Rows.Count > 3; i++) tbl.Rows[2].Delete();
                //Dynamically adds row in Word Table
                for (int i = 1; i < stockTable.Count - 1; i++) tbl.Rows.Add();
                int rowCounter = 2;

                for (int i = 0; i < stockTable.Count - 1; i++)
                {


                    tbl.Cell(rowCounter, 1).Range.Text = table[i][0];
                    tbl.Cell(rowCounter, 2).Range.Text = table[i][1];
                    tbl.Cell(rowCounter, 3).Range.Text = table[i][2];
                    tbl.Cell(rowCounter, 4).Range.Text = table[i][3];
                    tbl.Cell(rowCounter, 5).Range.Text = table[i][4];
                    if (i == 0)
                    {
                        tbl.Cell(rowCounter, 6).Range.Text = (((decimal.Parse(table[i][2]) - decimal.Parse("0.00")) / (decimal.Parse(table[i][2]))) * 100).ToString("0.00");
                    }
                    else
                    {
                        tbl.Cell(rowCounter, 6).Range.Text = (((decimal.Parse(table[i][2]) - decimal.Parse(table[i - 1][2])) / (decimal.Parse(table[i][2]))) * 100).ToString("0.00");
                    }

                    if (rowCounter < stockTable.Count - 1) ws1.Rows[rowCounter].Insert();

                    ws1.Cells[rowCounter, "A"].Value = table[i][0]; Thread.Sleep(25);
                    ws1.Cells[rowCounter, "B"].Value = table[i][2].ToString(); Thread.Sleep(25);
                    //Dynamically adds row in WorkBook
                    if (rowCounter < stockTable.Count - 1) ws2.Rows[rowCounter].Insert();
                    ws2.Cells[rowCounter, "A"].Value = table[i][0]; Thread.Sleep(25);
                    ws2.Cells[rowCounter, "B"].Value = table[i][2].ToString(); Thread.Sleep(25);
                    rowCounter++; if (rowCounter == stockTable.Count + 1) break;
                }

                return "Done";
            }
            catch (Exception ex)
            {
                Log.Error("Couldn't bind data in chart.", ex);
                throw new Exception("Couldn't bind data in chart.", ex);
            }
            finally
            {
                wb1.Close(true);

                if (shape1 != null) Marshal.ReleaseComObject(shape1);
                if (ws1 != null) Marshal.ReleaseComObject(ws1);
                if (wb1 != null) Marshal.ReleaseComObject(wb1);
                ws1 = null; wb1 = null;
                wb2.Close(true);

                if (shape2 != null) Marshal.ReleaseComObject(shape2);
                if (ws2 != null) Marshal.ReleaseComObject(ws2);
                if (wb2 != null) Marshal.ReleaseComObject(wb2);
                ws2 = null; wb2 = null;
                connection.Close();
            }
        }
        public string SummaryReportGenerator(Word.Document oWordDoc, string monthYr)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;


            ContentControl contentMonthYear = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark1].Range.ContentControls[1];
            Word.Table summaryTable1 = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark2].Range.Tables[1];
            Word.Table summaryTable2 = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark3].Range.Tables[1];
            ContentControl comments = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark4].Range.ContentControls[1];
            ContentControl contentNetWorth = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark5].Range.ContentControls[1];
            Word.Table topProfitableTable = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark6].Range.Tables[1];
            Word.Table nonProfitableTable = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark7].Range.Tables[1];
            ContentControl contentPieChart = oWordDoc.Bookmarks[ContentControl_Repository.oBookMark8].Range.ContentControls[1];

            Log.Verbose("Fetching data for binding contents...");

            decimal totalQuantity = NetQuantity(month);
            if (totalQuantity == 0) Log.Warning("Net Worth is null here.");

            List<string> profitCompanies = ProfitableCompanies(month);
            if (profitCompanies == null) Log.Warning($"Profitable Companies for {month} is null here.");

            List<string> nonProfitCompanies = NonProfitableCompanies(month);
            if (nonProfitCompanies == null) Log.Warning($"Non Profitable Companies for {month} is null here.");

            List<dynamic> summaryReportDetails1 = GetSummaryReportDetails1(month);
            List<dynamic> summaryReportDetails2 = GetSummaryReportDetails2(month);
            if (summaryReportDetails1 == null || summaryReportDetails1 == null) Log.Warning($"Summary Report for {month} is null here.");


            Word.InlineShape shape = contentPieChart.Range.InlineShapes[1];
            Word.Chart chart = shape.Chart;
            Excel.Workbook? wb = chart.ChartData.Workbook;
            Excel.Worksheet? ws = wb.Worksheets["Sheet1"];

            Log.Information("Binding data in Word template...");
            contentMonthYear.Range.Text = monthYr;
            contentNetWorth.Range.Text = totalQuantity.ToString();
            Log.Information("Data binded in Content Control.");

            int rowCounter = 2;
            foreach (var row in profitCompanies)
            {
                topProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();
                topProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();
                topProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();

                rowCounter++;
            }

            rowCounter = 2;
            foreach (var row in nonProfitCompanies)
            {
                nonProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();
                nonProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();
                nonProfitableTable.Cell(rowCounter, 1).Range.Text = row.ToString();

                rowCounter++;
            }

            try
            {
                //Deletes the rows after 1st company until no.of rows = 3 including Heading
                if (summaryTable1.Rows.Count > 3) for (int i = 1; summaryTable1.Rows.Count > 3; i++) summaryTable1.Rows[2].Delete();
                //Dynamically adds row in Word Table
                for (int i = 1; i < summaryReportDetails1.Count; i++) summaryTable1.Rows.Add();
                rowCounter = 2;
                foreach (var row in summaryReportDetails1)
                {
                    summaryTable1.Cell(rowCounter, 1).Range.Text = row.Company_Ticker.Trim();
                    summaryTable1.Cell(rowCounter, 2).Range.Text = row.Total_Quantity.ToString();
                    summaryTable1.Cell(rowCounter, 3).Range.Text = row.Average_Cost.ToString("0.00");
                    summaryTable1.Cell(rowCounter, 4).Range.Text = row.Total_Investment.ToString("0.00");
                    summaryTable1.Cell(rowCounter, 5).Range.Text = row.Current_Price.ToString("0.00");
                    summaryTable1.Cell(rowCounter, 6).Range.Text = row.Total_Amount.ToString("0.00");
                    summaryTable1.Cell(rowCounter, 7).Range.Text = row.Profit_Loss.ToString("0.00");

                    //Dynamically adds row in WorkBook
                    if (rowCounter < summaryReportDetails1.Count) ws.Rows[rowCounter].Insert();

                    ws.Cells[rowCounter, "A"].Value = row.Company_Ticker.Trim(); Thread.Sleep(25);
                    ws.Cells[rowCounter, "B"].Value = row.Total_Amount.ToString(); Thread.Sleep(25);

                    rowCounter++;
                }

                //Table 2
                //Deletes the rows after 1st company until no.of rows = 3 including Heading
                if (summaryTable2.Rows.Count > 3) for (int i = 1; summaryTable2.Rows.Count > 3; i++) summaryTable2.Rows[2].Delete();
                //Dynamically adds row in Word Table
                for (int i = 1; i < summaryReportDetails2.Count; i++) summaryTable2.Rows.Add();
                rowCounter = 2;
                foreach (var row in summaryReportDetails2)
                {
                    summaryTable2.Cell(rowCounter, 1).Range.Text = row.sector;
                    summaryTable2.Cell(rowCounter, 2).Range.Text = row.total_investment;
                    summaryTable2.Cell(rowCounter, 3).Range.Text = row.profitLoss;
                    summaryTable2.Cell(rowCounter, 4).Range.Text = row.percentage;

                    rowCounter++;
                }
                string summary;
                var comment = (from RT in _context.PrbReportTypes where RT.Month == monthYr && RT.ReportTypeCode == "S" select new { RT.Commentary }).FirstOrDefault();


                if (comment == null) summary = " ";
                else
                {
                    summary = comment.Commentary;
                    summary = Regex.Replace(summary, @"<(br|BR)\s{0,1}\/{0,1}>", Environment.NewLine);
                    summary = Regex.Replace(summary, @"(\</?p(.*?)/?\>)", string.Empty);
                    //summary = Regex.Replace(summary, @"(\</?p(.*?)/?\>)", string.Empty);
                }
                comments.Range.Text = summary;



                RichTextRender(comments, summary);
            }
            catch (COMException ex)
            {
                Log.Error("Couldn't bind data in chart.", ex);
                throw new Exception("Couldn't bind data in chart.", ex);
            }
            finally
            {
                wb.Close(true);

                Log.Debug($"Saving the document under directory {appSettings.ReportPath?.SummaryReportPath + ".docx..."}");
                oWordDoc.SaveAs2(appSettings.ReportPath?.SummaryReportPath + ".docx");
                Log.Verbose("Exporting the document as PDF format...");
                oWordDoc.ExportAsFixedFormat(appSettings.ReportPath?.SummaryReportPath + ".pdf", WdExportFormat.wdExportFormatPDF);


                if (shape != null) Marshal.ReleaseComObject(shape);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null) Marshal.ReleaseComObject(wb);
                ws = null; wb = null;
                connection.Close();
            }
            return null;
        }
        public void RichTextRender(ContentControl comments, string summary)
        {
            try
            {
               

                comments.Range.Text = summary;



                System.Span<char> content = new char[2000];
                comments.Range.Text.CopyTo(content);
                var regex = new Regex(@"(\</?(.*?)/?\>)");
                var matcher = regex.Match(content.ToString());



                while (content.ToString() != HttpUtility.HtmlEncode(content.ToString()))
                {
                    //RegEx for get tag (which tag). So that it can be used in HTML Agility
                    regex = new Regex(@"(\</?(.*?)/?\>)");
                    if (matcher.Length == 0) break;





                    HtmlDocument doc = new HtmlDocument();
                    doc.LoadHtml(content.ToString());
                    string match = matcher.Groups[2].Value;
                    if (matcher.Groups[2].Value.Length > 3) if (matcher.Groups[2].Value.Substring(0, 4) == "span") match = "span";
                    var node = doc.DocumentNode.SelectNodes($"//{match}").FirstOrDefault();



                    //For Formatting the Inline characters
                    Word.Range rn = comments.Range;
                    rn.Start = node.InnerStartIndex + 1;
                    rn.End = node.InnerStartIndex + node.InnerLength + 1;



                    switch (match)
                    {
                        case "span":
                            if (matcher.Groups[2].Value.Substring(15, 4) == "font")
                            {
                                if (matcher.Groups[2].Value.Substring(20, 4) == "mono") rn.Font.Name = "Courier New";
                            }
                            if (matcher.Groups[2].Value.Substring(15, 4) == "size")
                            {
                                if (matcher.Groups[2].Value.Substring(20, 4) == "huge") rn.Font.Size = (float)24.5;
                                else if (matcher.Groups[2].Value.Substring(20, 5) == "large") rn.Font.Size = (float)14.5;
                                else if (matcher.Groups[2].Value.Substring(20, 5) == "small") rn.Font.Size = (float)7.5;
                            }
                            break;
                        case "h1": rn.Font.Size = 24; break;
                        case "h2": rn.Font.Size = 18; break;
                        case "h3": rn.Font.Size = 14; break;
                        case "h4": rn.Font.Size = 10; break;
                        case "strong": rn.Font.Bold = 1; break;
                        case "sup": rn.Font.Superscript = 1; break;
                        case "u": rn.Font.Underline = Word.WdUnderline.wdUnderlineSingle; break;
                        case "em": rn.Font.Italic = 1; break;
                        case "s": rn.Font.StrikeThrough = 1; break;
                    }
                    //For Removing Opening Tag
                    Word.Range deleteRange1 = comments.Range;
                    deleteRange1.Start = node.OuterStartIndex + 1;
                    deleteRange1.End = node.InnerStartIndex + 1; // Returns end of opening tag
                    deleteRange1.Delete();
                    comments.Range.Text.CopyTo(content);
                    content = new char[content.ToString().Length - (deleteRange1.End - deleteRange1.Start)];
                    //For Removing Closing Tag
                    comments.Range.Text.CopyTo(content);



                    Word.Range deleteRange2 = comments.Range;
                    deleteRange2.Start = node.InnerStartIndex + node.InnerLength + 1 - (node.InnerStartIndex - node.OuterStartIndex);
                    deleteRange2.End = deleteRange2.Start + match.Length + 3;
                    deleteRange2.Delete();
                    content = new char[content.ToString().Length - (deleteRange2.End - deleteRange2.Start)];
                    comments.Range.Text.CopyTo(content);
                    matcher = regex.Match(content.ToString());
                }



            }
            catch (COMException ex)
            {
                Log.Error("Couldn't bind data.", ex);
                throw new Exception("Couldn't bind data.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        public string DisclaimerReportGenerator(Word.Document oWordDoc, string monthYr)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;

            object oBookMark1 = "Disclaimer";
            ContentControl contentDisclaimer = oWordDoc.Bookmarks[oBookMark1].Range.ContentControls[1];

            Log.Verbose("Fetching commentary from database...");
            string disclaimer;

            var disclaimerDetails = (from RT in _context.PrbReportTypes where RT.Month == monthYr && RT.ReportTypeCode == "D" select new { RT.Commentary }).FirstOrDefault();

            try
            {
                Log.Information("Binding data in Word template...");
                if (disclaimerDetails != null) disclaimer = disclaimerDetails.Commentary;
                else disclaimer = DisclaimerFetcher(oWordDoc, monthYr);
                disclaimer = Regex.Replace(disclaimer, @"<(br|BR)\s{0,1}\/{0,1}>", Environment.NewLine);
                disclaimer = Regex.Replace(disclaimer, @"(\</?p(.*?)/?\>)", string.Empty);
                contentDisclaimer.Range.Text = disclaimer;
                RichTextRender(contentDisclaimer, disclaimer);
                contentDisclaimer.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                Log.Information("Data binded in Content Control");
            }
            catch (COMException ex)
            {
                Log.Error("Couldn't bind data in chart.");
                throw new Exception("Couldn't bind data in chart.", ex);
            }
            finally
            {
                Log.Debug($"Saving the document under directory {appSettings.ReportPath?.DisclaimerReportPath + ".docx..."}");
                oWordDoc.SaveAs2(appSettings.ReportPath?.DisclaimerReportPath + ".docx");
                Log.Verbose("Exporting the document as PDF format...");
                oWordDoc.ExportAsFixedFormat(appSettings.ReportPath?.DisclaimerReportPath + ".pdf", WdExportFormat.wdExportFormatPDF);
                connection.Close();
            }
            return null;
        }
        public string SummaryFetcher(Word.Document oWordDoc, string monthYr)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;



            object oBookMark1 = "Comments";
            ContentControl contentSummary = oWordDoc.Bookmarks[oBookMark1].Range.ContentControls[1];
            System.Span<char> summary = new char[2000];
            var summaryDetails = (from RT in _context.PrbReportTypes where RT.Month == monthYr && RT.ReportTypeCode == "S" select new { RT.Commentary }).FirstOrDefault();
            try
            {
                Log.Information("fetch from Word template...");
                if (summaryDetails != null) return summaryDetails.Commentary;
                else return null;
                //{
                //    contentSummary.Range.Copy();
                //    contentSummary.Range.Text.CopyTo(summary);

                //    Log.Information("Data fetch in Content Control");
                //    return summary.ToString();
                //}
            }
            catch (COMException ex)
            {
                Log.Error("Couldn't fetch data in chart.");
                throw new Exception("Couldn't fetch data in chart.", ex);
            }
        }
        public string DisclaimerFetcher(Word.Document oWordDoc, string monthYr)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;



            object oBookMark1 = "Disclaimer";
            ContentControl contentDisclaimer = oWordDoc.Bookmarks[oBookMark1].Range.ContentControls[1];
            System.Span<char> disclaimer = new char[2000];
            var disclaimerDetails = (from RT in _context.PrbReportTypes where RT.Month == monthYr && RT.ReportTypeCode == "D" select new { RT.Commentary }).FirstOrDefault();
            try
            {
                Log.Information("fetch from Word template...");
                if (disclaimerDetails != null) return disclaimerDetails.Commentary;
                else
                {
                    contentDisclaimer.Range.Copy();
                    contentDisclaimer.Range.Text.CopyTo(disclaimer);

                    Log.Information("Data fetch in Content Control");
                    return disclaimer.ToString();
                }
            }
            catch (COMException ex)
            {
                Log.Error("Couldn't fetch data in chart.");
                throw new Exception("Couldn't fetch data in chart.", ex);
            }
        }
        public string Generate(string report)
        {
            try
            {
                string data;
                string workFlow = "AutomationWorkFlow";
                data = report;
                Log.Information("Initiating Rule Engine Executor...");
                var str = this.rulesExecutor.GetHomeEngine(data, workFlow);
                return str.Result.ToString();
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                connection.Close();
            }
        }

        public IDictionary<string, string> FetchUserName(string mailId)
        {
            IDictionary<string, string> userDetails = new Dictionary<string, string>();
            try
            {
                var user = (from U in _context.PrbUsers
                            join RC in _context.PrbRoleCodes
                            on U.RoleCode equals RC.RoleCode
                            where U.UserMailId == mailId
                            select new { U.UserName, RC.RoleDesc }).ToList();
                if (user == null) Log.Warning($"There is no User with {mailId}.");
                user.ForEach(x =>
                {
                    userDetails.Add(x.UserName, x.RoleDesc);
                });
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                connection.Close();
            }
            return userDetails;
        }

        //Getting Ticker
        public List<string> GettingTicker(string monthYr)
        {
            try
            {
                int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
                int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;
                List<string> tickerlists = new List<string>();
                Log.Verbose($"Fetching all TICKERS from Database for {monthYr}.");
                var tickers = (from P in _context.PrbHoldingDetails where P.TransactionDate.Month == month && P.TransactionDate.Year == year select new { ticker = P.CompanyTicker.Trim() }).Distinct().ToList();
                if (tickers.Count == 0) Log.Warning("Database doesn't have the Company to fetch.");
                tickers.ForEach(x =>
                {
                    tickerlists.Add(x.ticker);
                });
                Log.Information($"Fetched all TICKERS from Database for {monthYr}.\n");
                return tickerlists;
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                connection.Close();
            }
        }
        //Ticker Synopsis
        public string TickerSynopsis(string tickerName)
        {
            try
            {
                var synopsis = (from p in _context.PrbTickers where p.CompanyTicker == tickerName select new { p.CompanyDesc }).Single();
                if (synopsis.CompanyDesc == "") Log.Warning("Database doesn't have the Company to fetch.");
                return synopsis.CompanyDesc.ToString();
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        public List<dynamic> StockDetails(int month, int year, string ticker)
        {
            try

            {

                string sql = "StockDetails";

                SqlConnection con = new SqlConnection(connectionStr.ToString());

                var param = new DynamicParameters();

                param.Add("@Company_Ticker", ticker);

                param.Add("@year", year);

                param.Add("@month", month);

                Log.Information("Getting holding Details.\n");

                var DetailedData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure);

                var detailedReport = ((IEnumerable)DetailedData).Cast<object>().ToList();

                return detailedReport;



            }

            catch (SqlException ex)

            {

                Log.Error("Something went wrong in Database.", ex);

                throw new Exception("Something went wrong in Database.", ex);

            }

            catch (Exception ex)

            {

                Log.Error("Something went wrong !", ex);

                throw new Exception("Something went wrong !", ex);

            }

            finally

            {

                connection.Close();

            }



        }
        public List<dynamic> OCBalance(string companyName, string monthYr)
        {



            Log.Verbose("Fetching Opening and closing...");
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;



            try
            {
                string sql = "currentOpening";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@Company_Ticker", companyName);
                param.Add("@month", month);
                Log.Information("Getting holding Details.\n");
                var balance = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure);
                List<dynamic> currentOpeningClosing = new List<dynamic>();
                var detailedReport = ((IEnumerable)balance).Cast<object>().ToList();
                detailedReport.ForEach(x =>
                {
                    currentOpeningClosing.Add((x));
                });
                Log.Information("Opening and closing balance fetched.\n");
                return currentOpeningClosing;
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                connection.Close();
            }
        }

        //Summary //Stored Procedure
        //Net Quantity
        public decimal NetQuantity(int month)
        {
            try
            {
                string sql = "NetQuantity";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@month", month);
                var SummaryData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure).ToList();
                if (SummaryData.Count == 0) Log.Warning("Net Quantity is null here.");
                decimal sum = 0;
                SummaryData.ForEach(x =>
                {
                    sum = sum + x.Net_Quantity;
                });
                return sum;
            }
            catch (SqlException ex)
            {
                Log.Error("Couldn't get values from Database");
                throw new Exception("Couldn't get values from Database", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        //Profitable Companies
        public List<string> ProfitableCompanies(int month)
        {
            try
            {
                List<string> ProfitCompanies = new List<string>();
                string sql = "ProfitCompanies";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@month", month);
                var SummaryData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure).ToList();
                if (SummaryData.Count == 0) Log.Warning("Profitable Companies is null here.");
                SummaryData.ForEach(x =>
                {
                    ProfitCompanies.Add(x.Company_Ticker.Trim());
                });
                return ProfitCompanies;
            }
            catch (SqlException ex)
            {
                Log.Error("Couldn't get values from Database");
                throw new Exception("Couldn't get values from Database", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        //Non Profitable Companies
        public List<string> NonProfitableCompanies(int month)
        {
            try
            {
                List<string> NonProfitCompanies = new List<string>();
                string sql = "NonProfitCompanies";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@month", month);
                var SummaryData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure).ToList();
                if (SummaryData.Count == 0) Log.Warning("Non Profitable Companies is null here.");
                SummaryData.ForEach(x =>
                {
                    NonProfitCompanies.Add((x.Company_Ticker).Trim());
                });
                return NonProfitCompanies;
            }
            catch (SqlException ex)
            {
                Log.Error("Couldn't get values from Database");
                throw new Exception("Couldn't get values from Database", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        //Summary Table 1
        public List<dynamic> GetSummaryReportDetails1(int month)
        {
            try
            {
                List<dynamic> summaryReportDetails1 = new List<dynamic>();
                string sql = "summarytable1";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@month", month);
                var SummaryData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure).ToList();
                SummaryData.ForEach(x =>
                {
                    summaryReportDetails1.Add((x));
                });
                return SummaryData;
            }
            catch (SqlException ex)
            {
                Log.Error("Couldn't get values from Database");
                throw new Exception("Couldn't get values from Database", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        //Summary Table 2
        public List<dynamic> GetSummaryReportDetails2(int month)
        {
            try
            {
                List<dynamic> summaryReportDetails2 = new List<dynamic>();
                string sql = "summarytable2";
                SqlConnection con = new SqlConnection(connectionStr.ToString());
                var param = new DynamicParameters();
                param.Add("@month", month);
                var SummaryData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure).ToList();
                var data = from s in SummaryData
                           group s by s.Sector_Name into grp
                           select new
                           {
                               sector = grp.Key,
                               total_investment = grp.Sum(tv => (double)tv.Total_Investment).ToString(),
                               profitLoss = (grp.Sum(pl => (double)pl.profitLoss) > 0) ? "Profit" : "Loss",
                               percentage = (Math.Abs((grp.Sum(p => (double)p.profitLoss) / grp.Sum(p => (double)p.Total_Investment)) * 100)).ToString("0.00")
                           };
                List<dynamic> summaryTable2 = data.Cast<object>().ToList();
                //summaryReportDetails2.Add(data);
                //List<dynamic> summaryTable2 = (List < dynamic >) summaryReportDetails2[0];
                return summaryTable2;
            }
            catch (SqlException ex)
            {
                Log.Error("Couldn't get values from Database");
                throw new Exception("Couldn't get values from Database", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        //Merge Document
        public string MergeDoc(Word.Application oWord, string monthYr, string password)
        {
            //object? mergedFilePassword = appSettings.Passwords?.MergedFilePassword;
            object? mergedFilePassword = password;

            List<string> allCompanyFilePaths = new List<string>();
            Log.Verbose("Fetching all DOCX file paths...");
            string[] allFilePaths = Directory.GetFiles(path: appSettings.ReportPath?.Path, "*.docx");
            for (int i = 0; i < allFilePaths.Length; i++)
            {
                if (allFilePaths[i] == appSettings.ReportPath?.SummaryReportPath + ".docx"
                    || allFilePaths[i] == appSettings.ReportPath?.DisclaimerReportPath + ".docx"
                    || allFilePaths[i] == appSettings.ReportPath?.MergedReportPath + "Merged.docx") continue;
                allCompanyFilePaths.Add(allFilePaths[i]);
            }
            allCompanyFilePaths.Insert(0, appSettings.ReportPath?.SummaryReportPath + ".docx");
            allCompanyFilePaths.Add(appSettings.ReportPath?.DisclaimerReportPath + ".docx");

            object missing = System.Type.Missing;
            object pageBreak = Word.WdBreakType.wdPageBreak;
            object sectionBreak = Word.WdBreakType.wdSectionBreakNextPage;

            object oMerge = allCompanyFilePaths[allCompanyFilePaths.Count - 1];
            Word.Document oWordDoc = new Word.Document();
            oWord.Visible = false;
            oWordDoc = oWord.Documents.Add(ref oMerge);
            Word.Selection selection = oWord.Selection;
            oWordDoc.PageSetup.TopMargin = 70.9f;

            try
            {
                String insertFile = "";
                Log.Information("Merging all DOCX file ...");
                for (int i = 0; i < allCompanyFilePaths.Count - 1; i++)
                {
                    insertFile = allCompanyFilePaths[i];
                    selection.InsertFile(insertFile, ref missing, ref missing, ref missing, ref missing);
                    selection.InsertBreak(ref sectionBreak);
                    Log.Debug($"{Path.GetFileName(allCompanyFilePaths[i])} is merged into Document.");
                }
            }
            catch (COMException ex)
            {
                Log.Error("Couldn't merge the documents.", ex);
                throw new Exception("Couldn't merge the documents.", ex);
            }
            finally
            {
                object enforceStyleLock = false;
                oWordDoc.Protect(WdProtectionType.wdAllowOnlyReading, ref missing, ref mergedFilePassword, ref missing, ref enforceStyleLock);
                oWordDoc.ReadOnlyRecommended = false;
                Log.Information("Merged document is Protected with password");

                Log.Debug($"Saving the document under directory {appSettings.ReportPath?.MergedReportPath + ".docx..."}");
                oWordDoc.SaveAs2(appSettings.ReportPath?.MergedReportPath + "Merged.docx");
                Log.Debug("Exporting the document as PDF format...");
                oWordDoc.ExportAsFixedFormat(appSettings.ReportPath?.MergedReportPath + "Merged.pdf", WdExportFormat.wdExportFormatPDF);

            }
            //FileInfo file = new FileInfo(WebRootPath + "\\Report\\Merged.docx");
            //if (file.Exists)//check file exsit or not  
            //{
            //    file.Delete();
            //}
            return "Done";
        }
        public string TemplateIdProvider(string companyticker)
        {
            try
            {
                var templateId = (from p in _context.PrbTickers where p.CompanyTicker == companyticker select new { p.TemplateId }).Single();

                return templateId.TemplateId.ToString();
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong.", ex);
                throw new Exception("Something went wrong.", ex);
            }
            finally
            {
                connection.Close();
            }
        }

    }
}
