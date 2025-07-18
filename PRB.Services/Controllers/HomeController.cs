using Dapper;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using PRB.Repository.DataContext;
using PRB.Repository.Repository;
using PRB.Domain.Model1;
using System.IO.Compression;
using PRB.Repository.Automation_Repository;
using System.Diagnostics;

namespace PRB.Services.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    [EnableCors("Mypolicy")]
    public class HomeController: ControllerBase
    {
        private readonly IConfiguration _config; 
        public readonly PRB_DB_Context _context;
        private readonly IPRB_Repository _IPRB_Repository;
        public static IWebHostEnvironment? _webHostEnvironment;
        public readonly Object WebRootPath;
        private string connectionStr = string.Empty;
        private readonly AppSettings appSettings;

        protected readonly IAutomation_Repository _IAutomation_Repository;
        public HomeController(IConfiguration config, PRB_DB_Context context, IPRB_Repository prb_Repository, IWebHostEnvironment webHostEnvironment, IAutomation_Repository iAutomation_Repository) : base()
        {

            _config = config;
            _context = context;
            this._IPRB_Repository = prb_Repository;
            _webHostEnvironment = webHostEnvironment;
            WebRootPath = _webHostEnvironment.WebRootPath ;
            connectionStr = config.GetConnectionString("MyDBConnection");
            appSettings = config.Get<AppSettings>();
            _IAutomation_Repository = iAutomation_Repository;
    }


        //[HttpGet("Automate")]
        //public object automate()
        //{
        //    string mailId = "raj@gmail.com";
        //    string monthYear = "Mar 2023";
        //    string file = "ACN";
        //    int templateId = 2;
        //    foreach (var process in Process.GetProcessesByName("WINWORD"))
        //    {
        //        process.Kill();
        //    }

        //    return _IPRB_Repository.automate(mailId, monthYear, file, templateId);
        //}
        [HttpGet("Automate/{mailId}&{monthYear}&{file}&{id}")]
        public object Automate(string mailId, string monthYear, string file,int id)
        {

            foreach (var process in Process.GetProcessesByName("WINWORD"))
            {
                process.Kill();
            }

            //int templateId = 0;
            return _IPRB_Repository.Automate(mailId, monthYear, file, id);
        }
        //Login Checking
        [HttpGet("loginChecking/{mailId}&{password}")]
        public object LoginChecking([FromRoute] string mailId, [FromRoute] string password)
        {
            return _IPRB_Repository.LoginChecking(mailId, password);
        }
        [HttpGet("Updatetemplate/{companyTickers}&{templateid}")]
        public string Updatetemplate(string companyTickers, int templateid)
        {
            return _IPRB_Repository.Updatetemplate(companyTickers, templateid);
        }
        //Get all reports Name 
        [HttpGet("gettingReports/{role}")]
    public object GettingReports(string role)
    {
      return _IPRB_Repository.GettingReports(role);
    }

    //Get all Company Ticker
    [HttpGet("gettingTicker/{monthYear}")]
        public object GettingTicker( string monthYear)
        {
            return _IPRB_Repository.GettingTicker(monthYear);
        }

        //Microsoft Rule Engine
        [HttpGet("getStatus/{Role}&{ReportStatus}&{ReportCommentary}")]
        public Task<string> GetStatus([FromRoute] string Role, [FromRoute] string ReportStatus, [FromRoute] string ReportCommentary)
        {
            return _IPRB_Repository.GetStatus(Role, ReportStatus, ReportCommentary);
        }

        //Checking if Author or Reviewer 1 is rejected or not for getting comments
        [HttpGet("commentsChecker/{monthYear}")]
        public object CommentsChecker(string monthYear)
        {
            return _IPRB_Repository.CommentsChecker(monthYear);
        }

        //Get all Company file path and download file
        [HttpGet("DownloadPDF/{fileName}")]
        public async Task<IActionResult> DownloadPDF(string fileName)
        {
            var filePath = (_IPRB_Repository.PathProvider(fileName)).Values.FirstOrDefault();
            var memory = new MemoryStream();
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                await stream.CopyToAsync(memory);
                memory.Position = 0;
                return File(memory, GetContentType(filePath));
            }
        }

        protected string GetContentType(string path)
        {
            var types = GetMimeTypes();
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return types[ext];
        }
        protected Dictionary<string, string> GetMimeTypes()
        {
            return new Dictionary<string, string>
                 {
                    {".txt", "text/plain"},
                    {".pdf", "application/pdf"},
                    {".doc", "application/vnd.ms-word"},
                    {".docx", "application/vnd.ms-word"},
                    {".xls", "application/vnd.ms-excel"},
                    {".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
                    {".png", "image/png"},
                    {".jpg", "image/jpeg"},
                    {".jpeg", "image/jpeg"},
                    {".gif", "image/gif"},
                    {".csv", "text/csv"}
                };

        }
        [HttpGet("TemplateProvider")]
        public object TemplateProvider()
        {
            return _IPRB_Repository.TemplateProvider();
        }

        //Detailed Report Data        
        [HttpGet("getHoldingDetails/{companyTickers}&{monthYr}")]
        public object GetHoldingDetails([FromRoute] string companyTickers, [FromRoute] string monthYr)
        {
            return _IPRB_Repository.GetHoldingDetails(companyTickers, monthYr);
        }

        //Checking entries oftable to store / update Summary and Disclaimer
        [HttpGet("checkingStatus/{filename}&{type}")]
        public string CheckingStatus(string filename, string type)
        {
            return _IPRB_Repository.CheckingStatus(filename, type);
        }

        //Storing Commentary for Disclaimer and Summary
        [HttpPost("storingCommentary")]
        public string StoringCommentary([FromBody] CommentaryModel details)
        {
            return _IPRB_Repository.StoringCommentary(details);
        }

        // Updating Commentary for Disclaimer and summary
        [HttpPut("UpdateCommentary")]
        public string UpdateCommentry([FromBody] CommentaryModel details)
        {
            return _IPRB_Repository.UpdateCommentary(details);
        }

        //Update Final Report Status
        [HttpPost("updateReportStatus")]
        public string UpdateReportStatus([FromBody]  UpdateReportStatus details)
        {
            return _IPRB_Repository.UpdateReportStatus(details);
        }


        //Update mutliple value rows
        [HttpPut("updateDetailedReportData")]
        public string UpdateDetailedReportData([FromBody] IEnumerable<UpdateDetailedReportDatas> data)
        {
            return _IPRB_Repository.UpdateDetailedReportData(data);
        }

        //Checking File status returns C
        [HttpGet("CheckingReport/{ReportMonth}")]
        public string CheckingReport([FromRoute] string ReportMonth)
        {
            return _IPRB_Repository.CheckingReport(ReportMonth);


        }



        [HttpGet("getTemplatesId")]
        public object getTemplatesId()
        {
            return _IPRB_Repository.getTemplatesId();
        }


        [HttpPut("updateTemplates/{companyTicker}&{t_Id}")]
        public string updateTemplates(string companyTicker, int t_Id)
        {
            return _IPRB_Repository.updateTemplates(companyTicker, t_Id);
        }

        [HttpGet("reportStatusRE/{filename}")]
    public string ReportStatusRE(string filename)
    {
      return _IPRB_Repository.ReportStatusRE(filename);
    }

    //Checking File status returns C
    [HttpGet("Auto")]
        public string Auto()
        {
            _IPRB_Repository.ReportPusher(DateTime.Now.AddMonths(-1).ToString("MMM yyyy"));

            return "No Discrepancy";
        }


        //Not for Project - BackUp Functions
        [HttpGet("Day-to-Day_BackUp")]
        public void zip()
        {
            string startPath = @"D:\PRB.Services";
            string zipPath = @"D:\PRB\Backup\PRB.Services_" + DateTime.Now.ToString("dd MMMM (HH-mm-ss)") + ".zip";
            if (!Directory.Exists("D:\\PRB_BackUp_Folder")) Directory.CreateDirectory("D:\\PRB_BackUp_Folder");

            if (System.IO.File.Exists("D:\\PRB_BackUp_Folder\\PRB.Services\\PRB.Services.sln")) Directory.Delete("D:\\PRB_BackUp_Folder\\PRB.Services", true);
            string fileName = "PRB.Services_" + DateTime.Now.ToString("dd MMMM");
            Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(startPath, "D:\\PRB_BackUp_Folder\\PRB.Services");
            new List<string>(Directory.GetFiles(@"D:\PRB\Backup")).ForEach(file => { if (file.ToUpper().Contains(fileName.ToUpper())) System.IO.File.Delete(file); });
            ZipFile.CreateFromDirectory("D:\\PRB_BackUp_Folder", zipPath);

            new List<string>(Directory.GetFiles(@"D:\")).ForEach(file => { if (file.ToUpper().Contains("PRB.Services__".ToUpper())) System.IO.File.Delete(file); });

        }
        [HttpGet("PerennialBackUp")]
        public void PerennialZip()
        {
            string startPath = @"D:\PRB.Services";
            string zipPath = @"D:\PRB.Services__" + DateTime.Now.ToString("dd MMMM (HH-mm-ss)") + ".zip";
            if (!Directory.Exists(path: "D:\\PRB_BackUp_Folder")) Directory.CreateDirectory("D:\\PRB_BackUp_Folder");

            if (System.IO.File.Exists("D:\\PRB_BackUp_Folder\\PRB.Services\\PRB.Services.sln")) Directory.Delete("D:\\PRB_BackUp_Folder\\PRB.Services", true);

            Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(startPath, "D:\\PRB_BackUp_Folder\\PRB.Services");
            ZipFile.CreateFromDirectory("D:\\PRB_BackUp_Folder", zipPath);

        }


    }

}

