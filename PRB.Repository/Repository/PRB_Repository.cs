
using Microsoft.Extensions.Configuration;
using PRB.Repository.DataContext;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using Microsoft.EntityFrameworkCore;
using Serilog;
using PRB.Services;
using Dapper;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using PRB.Domain.Model1;
using PRB.Repository.Automation_Repository;
using Nest;
using System.Security.Cryptography;
using System.Text;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Newtonsoft.Json.Linq;
using System.Text.Json;
using Newtonsoft.Json;
using System.Dynamic;
using Newtonsoft.Json.Converters;
using System.Reflection;
using RulesEngine.Models;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;

namespace PRB.Repository.Repository
{
    public interface IPRB_Repository
    {
        public object Automate(string mailId, string monthYear, string file, int templateId);
        public object LoginChecking(string mailId, string password);
    public List<string> GettingReports(string role);
        public List<string> GettingTicker(string monthYr);
        public IDictionary<string, string> PathProvider(string filename);
        public object TemplateProvider();
        public object GetHoldingDetails(string companyTickers, string monthYr);
        public string CheckingStatus(string filename, string type);
        public string StoringCommentary(CommentaryModel details);
        public string UpdateCommentary(CommentaryModel details);
        public string UpdateReportStatus(UpdateReportStatus details);
        public string RoleConverter(string reportStatus);
        public string UpdateDetailedReportData(IEnumerable<UpdateDetailedReportDatas> data);
        public string ReportStatusRE(string filename);
        public string CheckingReport(string ReportMonth);
        public object CommentsChecker(string monthYear);
        public Task<string> GetStatus(string Role, string ReportStatus, string ReportCommentary);
        public void ReportPusher(string monthYr);
        public string Updatetemplate(string companyTickers,int templateid);
        public string updateTemplates(string companyTicker, int t_Id);
        public object getTemplatesId();

    }
    public class PRBRepository : IPRB_Repository
    {
        private readonly IConfiguration _configuration;
        private readonly PRB_DB_Context _context;
        private string connectionStr = string.Empty;
        private readonly AppSettings appSettings;
        protected readonly SqlConnection connection;
        protected readonly IRuleExecutor rulesExecutor;
        protected readonly IAutomation_Repository _IAutomation_Repository;

        public PRBRepository(IConfiguration configuration, PRB_DB_Context context, IRuleExecutor rExecutor, IAutomation_Repository iAutomation_Repository)
        {
            this._configuration = configuration;
            this._context = context;
            connectionStr = configuration.GetConnectionString("MyDBConnection");
            this.connection = new SqlConnection(connectionStr);
            appSettings = configuration.Get<AppSettings>();
            this.rulesExecutor = rExecutor;
            _IAutomation_Repository = iAutomation_Repository;
        }
        //Word Automation
        public  object Automate(string mailId,string monthYear, string file, int templateId)
        {
            IDictionary<string, string> userDetails = _IAutomation_Repository.FetchUserName(mailId);
            Log.Information("---Automation Started---\nInitiating Word Application");           
            Word.Application oWord = new Word.Application();
            Word.Document oWordDoc = new Word.Document();

            string password;
            string flow=_IAutomation_Repository.Generate(file); 
            switch (flow)
            {
                case "1":
                    //Detailed Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Detailed Report...");
                    DetailedReportOrginator(oWord, oWordDoc, monthYear, file, templateId);

                    //Summary Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Summary Report...");
                    SummaryReportOrginator(oWord, oWordDoc, monthYear);

                    //Disclaimer Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Disclaimer Report...");
                    DisclaimerReportOrginator(oWord, oWordDoc, monthYear);
                    //Generating Passcode
                    password=Encrypt();
                    
                    //Merging all document
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Merging all docx...");
                    MergeReportOrginator(oWord, monthYear, password);
                    return "Completed";
                case "2":
                    //Decrypt the password
                    password = Decrypt();
                    //Detailed Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Detailed Report...");
                    DetailedReportOrginator(oWord, oWordDoc, monthYear, file, templateId);
                    //Summary Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Summary Report...");
                    SummaryReportOrginator(oWord, oWordDoc, monthYear);                    
                    //Merging all document
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Merging all docx...");
                    MergeReportOrginator(oWord, monthYear, password);
                    return "Completed";
                case "3":
                    //Decrypt the password
                    password = Decrypt();
                    //Summary Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Summary Report...");
                    SummaryReportOrginator(oWord, oWordDoc, monthYear);
                    //Merging all document
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Merging all docx...");
                    MergeReportOrginator(oWord, monthYear, password);
                    return "Completed";
                case "4":
                    //Decrypt the password
                    password = Decrypt();
                    //Disclaimer Report
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} generating Disclaimer Report...");
                    DisclaimerReportOrginator(oWord, oWordDoc, monthYear);
                    //Merging all document
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Merging all docx...");
                    MergeReportOrginator(oWord, monthYear,password);
                    return "Completed";
                case "5":
                    //Fetching Disclaimer Content
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Fetching Disclaimer content...");
                    Object? oTemplatePath = appSettings.TemplatePath?.DisclaimerReportTemplatePath;
                    oWordDoc = oWord.Documents.Add(ref oTemplatePath);
                    Log.Verbose("Template opened.");
                    object? templatePassword = appSettings.Passwords?.TemplatePassword;
                    oWordDoc.Unprotect(ref templatePassword);
                    return _IAutomation_Repository.DisclaimerFetcher(oWordDoc, monthYear);
                case "6": 
                    //Fetching Disclaimer Content
                    Log.Information($"{userDetails.First().Value} : {userDetails.First().Key} Fetching Summary content...");
                    Object? oTemplatePath1 = appSettings.TemplatePath?.SummaryReportTemplatePath;
                    oWordDoc = oWord.Documents.Add(ref oTemplatePath1);
                    Log.Verbose("Template opened.");
                    object? templatePassword1 = appSettings.Passwords?.TemplatePassword;
                    oWordDoc.Unprotect(ref templatePassword1);
                    return _IAutomation_Repository.SummaryFetcher(oWordDoc, monthYear);
                default: return "Not Completed";
            }
        }
        //For frontend services
        //Login Checking
        public object LoginChecking(string mailId, string password)
        {
            try
            {
                Log.Verbose("Fetching Username and Role Code... ");
                var user = (from us in _context.PrbUsers
                            join rc in _context.PrbRoleCodes on us.RoleCode equals rc.RoleCode
                            where us.UserMailId == mailId && us.Password == password
                            select new { rc.RoleDesc, us.UserName }).SingleOrDefault();
                if (user is null)
                {
                    Log.Warning($"No User with {mailId} and {password}");
                    return "User Name or Password Invalid ";
                }
                Log.Information($"{mailId} logging in.\n");
                return user;
            }
            catch (SqlException ex)
            {
                Log.Error("Something went wrong in Database.", ex);
                throw new Exception("Something went wrong in Database.", ex);
            }
            catch(Exception ex)
            {
                Log.Error("Something went wrong !",ex);
                throw new Exception("Something went wrong !",ex);
            }
            finally
            {
                connection.Close();
            }
        }
        //Jan 2023,Feb 2023
        public List<string> GettingReports(string role)
        {
            string monthYr = DateTime.Now.AddMonths(-1).ToString("MMM yyyy");
            try
            {
        List<string> reportmonths = new List<string>();
        List<string> reportMonths = new List<string>();
        Log.Verbose($"Fetching all Reports from Database.");
        var months = (from HD in _context.PrbHoldingDetails where HD.TransactionDate.Month < ((DateTime.Now.Month)) orderby (HD.TransactionDate) select new { month = Convert.ToDateTime(HD.TransactionDate).ToString("MMM") }).ToList();
        if (months.Count == 0) Log.Warning("Database doesn't have the data to fetch.");
        months.ForEach(x =>
        {
          reportmonths.Add(x.month);
        });
        for (int i = 0; i < reportmonths.Count; i++)
        {
          if (reportMonths.Contains(reportmonths[i] + " 2023"))
          {
            continue;
          }
          else
          {
            reportMonths.Add(reportmonths[i] + " 2023");
          }
        }



                var status = (from T in _context.PrbReportStatuses
                              where T.ReportMonth == monthYr && T.ReportStatusCode.Trim() == role
                              orderby T.ReportDate ascending
                              select new
                              {
                                  reportStatusCode = T.ReportStatusCode,

                              }).LastOrDefault();
                //status.ForEach(x =>
                //{
                //  entries.Add(x.reportStatusCode);
                //});


                if (role.Trim() != "D" && status is null ) reportMonths.RemoveAt(reportMonths.Count - 1);
                Log.Information($"Reports Fetched.\n");
        return reportMonths;
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

        //Final status
        public Task<string> GetStatus(string Role, string ReportStatus, string ReportCommentary)
        {
            try
            {
                var data = new MRERoleCode();
                string workFlow = "StageWorkFlow";
                data.Role = Role;
                data.ReportStatus = ReportStatus;
                data.ReportCommentry = ReportCommentary;
                Log.Information("Initiating Rule Engine Executor...");
                var str = this.rulesExecutor.GetHomeEngine(data, workFlow);
                return str;
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

        //Current Status of the Report
        public string ReportStatusRE(string filename)
        {
            try
            {
                var status = (from T in _context.PrbReportStatuses
                              where T.ReportMonth == filename
                              orderby T.ReportDate ascending
                              select new
                              {
                                  reportStatusCode = T.ReportStatusCode,

                              }).LastOrDefault();
                return status.reportStatusCode;
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
                tickers.ForEach(x => {
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
        public IDictionary<string, string> PathProvider(string filename)
        {
            IDictionary<string, string> fileNamePath = new Dictionary<string, string>();
            string[] allFilePaths ;
            Log.Verbose("Fetching all PDF file name and Paths...");
            if (filename.Any(Char.IsWhiteSpace))
            {
                filename = filename.Replace(" ", "_")+"_Merged";
                allFilePaths = Directory.GetFiles(path: appSettings.Records.Path, $"{filename}.pdf");                
            }

            else allFilePaths = Directory.GetFiles(path: appSettings.ReportPath?.Path, $"{filename}.pdf");

            for (int i = 0; i < allFilePaths.Length; i++) fileNamePath.Add(Path.GetFileNameWithoutExtension(allFilePaths[i]), allFilePaths[i]);

            return fileNamePath;
        }
        public object TemplateProvider()
        {
            try
            {
                var templateData = (from P in _context.PrbTemplatePaths
                                    select new
                                    {
                                        id = P.TemplateId,
                                        fileName = P.FileName,
                                        date = P.ExpiryDate.ToShortDateString()
                                    }).ToList();
                return templateData;
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
        public object GetHoldingDetails(string companyTickers, string monthYr)
        {
            int month = DateTime.ParseExact(monthYr[..3], "MMM", CultureInfo.CurrentCulture).Month;
            int year = DateTime.ParseExact(monthYr[4..], "yyyy", CultureInfo.CurrentCulture).Year;
            var DetailedData=_IAutomation_Repository.StockDetails(month, year, companyTickers);
            return DetailedData;
            //try
            //{
            //    string sql = "StockDetails";
            //    SqlConnection con = new SqlConnection(connectionStr.ToString());
            //    param.Add("@Company_Ticker", companyTickers);
            //    param.Add("@year", year);
            //    param.Add("@month", month);
            //    Log.Information("Getting holding Details.\n");
            //    var DetailedData = con.Query(sql, param, commandType: System.Data.CommandType.StoredProcedure);
            //    return DetailedData;
            //}
            //catch (SqlException ex)
            //{
            //    Log.Error("Something went wrong in Database.", ex);
            //    throw new Exception("Something went wrong in Database.", ex);
            //}
            //catch (Exception ex)
            //{
            //    Log.Error("Something went wrong !", ex);
            //    throw new Exception("Something went wrong !", ex);
            //}
            //finally
            //{
            //    connection.Close();
            //}
        }        
        //Check to Update or Store
        public string CheckingStatus(string filename, string type)
        {
            try
            {
                var status = _context.PrbReportTypes.Where(x => x.Month == filename && x.ReportTypeCode == type).FirstOrDefault();
                if (status != null) return "Match found";
                else return "No Match";
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

        public string StoringCommentary(CommentaryModel details)
        {
            try
            {
                var data = new PrbReportType();
                data.Month = details.Month;
                data.ReportTypeCode = details.ReportTypeCode;
                data.Commentary = details.Commentary;

                if(details.Commentary == null) Log.Warning("Commentary value is null. Commentary cannot be null.");
                _context.PrbReportTypes.Add(data);
                _context.SaveChanges();
                Log.Information("Commentary  stored in Database.");
                return "Successfully Stored";
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
        public string UpdateCommentary(CommentaryModel details)
        {
            var data = new PrbReportType();
            data.Month = details.Month;
            data.ReportTypeCode = details.ReportTypeCode;
            data.Commentary = details.Commentary;
            try
            {
                var updateCommentary = _context.PrbReportTypes.Where(x => x.Month == details.Month && x.ReportTypeCode == details.ReportTypeCode).Single();
                if (updateCommentary == null)
                {
                    Log.Warning($"Couldn't find data for {details.Month} and {details.ReportTypeCode} to update.");
                    return "Empty";
                }
                else
                {
                    updateCommentary.Commentary = data.Commentary;
                    _context.SaveChanges();
                    Log.Information("Commentary updated into Database.\n");
                    return "Updated";
                }
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

        public string Updatetemplate(string companyTickers, int templateid)  {
            try
            {
                var updateCommentary = _context.PrbTickers.Where(x => x.CompanyTicker == companyTickers).Single();
                if (updateCommentary == null)
                {
                    Log.Warning($"Couldn't find data for {updateCommentary.CompanyTicker}  to update.");
                    return "Empty";
                }
                else
                {
                    updateCommentary.TemplateId = templateid;
                    _context.SaveChanges();
                    Log.Information("Commentary updated into Database.\n");
                    return "Updated";
                }
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
            return null;

        }

        //Putting Entries while Report status changes
        public string UpdateReportStatus(UpdateReportStatus details)
        {
            try
            {
                //Entry in Table Report Status.
                Log.Information("Updating Report status...");
                var data = new PrbReportStatus();
                if (details.ReportMonth == null || details.ReportDate == null) Log.Warning("Report Month and Report Date shouldn't be null value.");
                data.ReportMonth = details.ReportMonth;
                data.ReportDate = (DateTime)details.ReportDate;
                data.RoleCode = details.RoleCode;
                data.ReportStatusCode = RoleConverter(details.ReportStatusCode);
                data.Comments = details.Comments;
                _context.PrbReportStatuses.Add(data);

                //Update Report Summary.
                var updateStatus = _context.PrbReportSummaries.Single(x => x.ReportMonth == data.ReportMonth);
                updateStatus.ReportStatusCode = data.ReportStatusCode;
                _context.SaveChanges();

                ReportPusher(DateTime.Now.AddMonths(-1).ToString("MMM yyyy"));

                Log.Information("Report updated into Database.\n");
                return "Successfully Stored";
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

        //Role Converter 
        public string RoleConverter(string reportStatus)
        {
      switch (reportStatus)
      {
        case "Draft": return "D ";
        case "Review-1": return "R1";
        case "Review-2": return "R2";
        case "Approver-1": return "A1";
        case "Approver-2": return "A2";
        case "Completed": return "C ";
        default: return null;
      }
    }

        //Updating Multiple values for Detailed Report
        public string UpdateDetailedReportData(IEnumerable<UpdateDetailedReportDatas> data)
        {
            try
            {
                foreach (var row in data)
                {
                    Log.Verbose("Updating Detailed report values...");
                    var updateTesting = _context.PrbHoldingDetails.Where(x => x.CompanyTicker == row.CompanyTicker && x.TransactionDate == row.TransactionDate).Single();
                    if (updateTesting == null)
                    {
                        Log.Warning($"{updateTesting} value is null here.");
                        return "Empty";
                    }
                    else
                    {
                        updateTesting.Quantity = row.Quantity;
                        updateTesting.Amount = row.Amount;
                        _context.SaveChanges();
                        Log.Information("Report updated into Database.\n");
                       
                    }

                }
                return "Updated";
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
            return "Successfull";
        }
       
        public string CheckingReport(string ReportMonth)
        {
            try
            {
                Log.Verbose("Fetching report status.");
                var response = (from p in _context.PrbReportStatuses where p.ReportMonth == ReportMonth orderby p.ReportDate descending select new { p.ReportStatusCode }).FirstOrDefault();
                if (response != null)
                {
                    Log.Information($"Report status{response.ReportStatusCode}.\n");
                    return response.ReportStatusCode;
                }
                else
                {
                    Log.Warning($"{response} value is null here.");
                    return "Generate";
                }
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
        public object CommentsChecker(string monthYear)
        {
            try
            {
                var status = (from T in _context.PrbReportStatuses
                              where T.ReportMonth == monthYear
                              orderby T.ReportDate ascending
                              select new
                              {
                                  roleCode = T.RoleCode,
                                  reportStatusCode = T.ReportStatusCode,
                                  comments = T.Comments
                              }).LastOrDefault();
     
        if (status == null || status.reportStatusCode == "D ") return status;
                else if (status.roleCode != "A " && status.reportStatusCode == "R1") return status;

                return null;
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


        public string DetailedReportOrginator(Word.Application oWord, Word.Document oWordDoc, string monthYear, string file, int templateId)
        {
            try
            {
                _IAutomation_Repository.DetaialedReportGenerator(oWord,oWordDoc, monthYear, file, templateId);
                Log.Information("Detailed Report generated.");

            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                Log.Verbose("Closing Word document.");
                oWordDoc.Close();
                Log.Information("Closing the Word Application");
                oWord.Quit();
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                Log.Information("Closing Word document.\n");
                oWordDoc.Close();
            }
            return null;
        }
        public string SummaryReportOrginator(Word.Application oWord, Word.Document oWordDoc, string monthYear)
        {
            try
            {
                Object? oTemplatePath = appSettings.TemplatePath?.SummaryReportTemplatePath;
                try
                {
                    oWordDoc = oWord.Documents.Add(ref oTemplatePath);
                    Log.Verbose("Template opened.");
                    object? templatePassword = appSettings.Passwords?.TemplatePassword;
                    oWordDoc.Unprotect(ref templatePassword);
                }
                catch (COMException ex)
                {
                    Log.Error($"Couldn't find the specified File Path {oTemplatePath}", ex);
                    throw new Exception("Sorry, we couldn't find your file. Was it moved, renamed, or deleted?", ex);
                }
                _IAutomation_Repository.SummaryReportGenerator(oWordDoc, monthYear);
                Log.Information("Summary Report generated.");
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                Log.Verbose("Closing Word document.");
                oWordDoc.Close();
                Log.Information("Closing the Word Application");
                oWord.Quit();
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                Log.Information("Closing Word document.\n");
                oWordDoc.Close();
            }
            return null;
        }
        public string DisclaimerReportOrginator(Word.Application oWord, Word.Document oWordDoc, string monthYear)
        {
            try
            {
                Object? oTemplatePath = appSettings.TemplatePath?.DisclaimerReportTemplatePath;
                try
                {
                    oWordDoc = oWord.Documents.Add(ref oTemplatePath);
                    Log.Verbose("Template opened.");
                    object? templatePassword = appSettings.Passwords?.TemplatePassword;
                    oWordDoc.Unprotect(ref templatePassword);
                }
                catch (COMException ex)
                {
                    Log.Error($"Couldn't find the specified File Path {oTemplatePath}", ex);
                    throw new Exception("Sorry, we couldn't find your file. Was it moved, renamed, or deleted?", ex);
                }
                _IAutomation_Repository.DisclaimerReportGenerator(oWordDoc, monthYear);
                Log.Information("Disclaimer Report generated.");
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong !", ex);
                Log.Verbose("Closing Word document.");
                oWordDoc.Close();
                Log.Information("Closing the Word Application");
                oWord.Quit();
                throw new Exception("Something went wrong !", ex);
            }
            finally
            {
                Log.Information("Closing Word document.\n");
                oWordDoc.Close();
            }
            return null;
        }
        public string MergeReportOrginator(Word.Application oWord, string monthYear, string password)
        {
            try
            {
                return _IAutomation_Repository.MergeDoc(oWord, monthYear, password);
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong.", ex);
                Log.Information("Closing the Word Application");
                oWord.Quit();
                throw new Exception("Something went wrong ", ex);
            }
            finally
            {
                Log.Information("Closing the Word Appliction");
                oWord.Quit();
                Log.Information("---Automation Terminated---\n");
            }
        }

        public string updateTemplates(string companyTicker, int t_Id)
        {
            try
            {
                var updateCommentary = _context.PrbTickers.Where(x => x.CompanyTicker == companyTicker).Single();
                if (updateCommentary == null)
                {
                    Log.Warning($"Couldn't find data for {updateCommentary.CompanyTicker} to update.");
                    return "Empty";
                }
                else
                {
                    updateCommentary.TemplateId = t_Id;
                    _context.SaveChanges();
                    Log.Information("Commentary updated into Database.\n");
                    return "Updated";
                }
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

        public object getTemplatesId()
        {

            try
            {
                var result = (from P in _context.PrbTickers
                              select new
                              {
                                  id = P.TemplateId,
                                  companyTicker = P.CompanyTicker.TrimEnd(),
                              }).ToList();


                return result;
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


        public string Encrypt()
        {
            try
            {
                string? passcode = PasswordGenerator();                
                string Return = null;
                string _key = "abcdefgh";
                string privatekey = "hgfedcba";
                byte[] privatekeyByte = { };
                privatekeyByte = Encoding.UTF8.GetBytes(privatekey);
                byte[] _keybyte = { };
                _keybyte = Encoding.UTF8.GetBytes(_key);
                byte[] inputtextbyteArray = System.Text.Encoding.UTF8.GetBytes(passcode);
                using (DESCryptoServiceProvider dsp = new DESCryptoServiceProvider())
                {
                    var memstr = new MemoryStream();
                    var crystr = new CryptoStream(memstr, dsp.CreateEncryptor(_keybyte, privatekeyByte), CryptoStreamMode.Write);
                    crystr.Write(inputtextbyteArray, 0, inputtextbyteArray.Length);
                    crystr.FlushFinalBlock();
                    string encrypted = Convert.ToBase64String(memstr.ToArray());

                    var data = new PrbReportSummary();
                    data.ReportMonth = DateTime.Now.AddMonths(-1).ToString("MMM yyyy");
                    data.ReportStatusCode = "D";
                    data.FilePassword = encrypted;
                    _context.PrbReportSummaries.Add(data);
                    _context.SaveChanges();

                    return encrypted;
                }
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong while encryption. !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        public string Decrypt()
        {
            try
            {
                string monthYr = DateTime.Now.AddMonths(-1).ToString("MMM yyyy");
                var code=(from P in _context.PrbReportSummaries where P.ReportMonth == monthYr select new { P.FilePassword }).FirstOrDefault();
                //DateTime.Now.AddMonths(-1).ToString("MMM yyyy");
                string encrptedCode = code.FilePassword;// appSettings.Passwords.EncryptedPassword;
                string x = null;
                string _key = "abcdefgh";
                string privatekey = "hgfedcba";
                byte[] privatekeyByte = { };
                privatekeyByte = Encoding.UTF8.GetBytes(privatekey);
                byte[] _keybyte = { };
                _keybyte = Encoding.UTF8.GetBytes(_key);
                byte[] inputtextbyteArray = new byte[encrptedCode.Replace(" ", "+").Length];
                //This technique reverses base64 encoding when it is received over the Internet.
                inputtextbyteArray = Convert.FromBase64String(encrptedCode.Replace(" ", "+"));
                using (DESCryptoServiceProvider dEsp = new DESCryptoServiceProvider())
                {
                    var memstr = new MemoryStream();
                    var crystr = new CryptoStream(memstr, dEsp.CreateDecryptor(_keybyte, privatekeyByte), CryptoStreamMode.Write);
                    crystr.Write(inputtextbyteArray, 0, inputtextbyteArray.Length);
                    crystr.FlushFinalBlock();
                    return Encoding.UTF8.GetString(memstr.ToArray());
                }
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong while decryption. !", ex);
                throw new Exception("Something went wrong !", ex);
            }
        }
        public string PasswordGenerator()
        {
            int length = 10;

            bool nonAlphanumeric = true;
            bool digit = true;
            bool lowercase = true;
            bool uppercase = true;

            StringBuilder password = new StringBuilder();
            Random random = new Random();
            try
            {
                while (password.Length < length)
                {
                    char c = (char)random.Next(32, 126);

                    password.Append(c);

                    if (char.IsDigit(c))
                        digit = false;
                    else if (char.IsLower(c))
                        lowercase = false;
                    else if (char.IsUpper(c))
                        uppercase = false;
                    else if (!char.IsLetterOrDigit(c))
                        nonAlphanumeric = false;
                }

                if (nonAlphanumeric)
                    password.Append((char)random.Next(33, 48));
                if (digit)
                    password.Append((char)random.Next(48, 58));
                if (lowercase)
                    password.Append((char)random.Next(97, 123));
                if (uppercase)
                    password.Append((char)random.Next(65, 91));
            }
            catch (Exception ex)
            {
                Log.Error("Something went wrong while generating password. !", ex);
                throw new Exception("Something went wrong !", ex);
            }
            return password.ToString();
        }

        public void ReportPusher(string monthYr)
        {
            var status = (from T in _context.PrbReportStatuses
                          where T.ReportMonth == monthYr
                          orderby T.ReportDate ascending
                          select new
                          {
                              reportStatusCode = T.ReportStatusCode
                          }).LastOrDefault();
            if (status is not null)
            if (status.reportStatusCode == "C ")
            {
                string filename = monthYr.Replace(" ", "_") + "_Merged";
                try
                {
                    Log.Information($"Moving files from {appSettings.ReportPath?.MergedReportPath} to {appSettings.Records?.Path}...");
                    System.IO.File.Move(appSettings.ReportPath?.MergedReportPath + "Merged.pdf"
                    , appSettings.Records?.Path + $"{filename}.pdf");
                }
                catch (FileNotFoundException ex)
                {
                    Log.Error("File not found !", ex);
                }
                catch (IOException ex)
                {
                    Log.Error("File already exists !", ex);
                }
                Directory.GetFiles(appSettings.ReportPath?.Path).ToList().ForEach(File.Delete);
                Log.Information($"Files deleted from {appSettings.ReportPath?.Path}.");
            }            
        }

    }
   
}
