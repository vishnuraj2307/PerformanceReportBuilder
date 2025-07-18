

namespace PRB.Services
{
    public class AppSettings
    {
        public ConnectionStrings? ConnectionStrings { get; set; }
        public string? RuleEnginePath { get; set; }
        public TemplatePath? TemplatePath { get; set; }
        public ReportPath? ReportPath { get; set; }
        public Records? Records { get; set; }
        public Images? Images { get; set; }
        public Passwords? Passwords { get; set; }
        public RichTextRender? RichTextRender { get; set; }
}
    public class ConnectionStrings
    {
        public string? MyDBConnection { get; set; }
        public string? TemplatePath { get;set; }
    }
    public class TemplatePath
    {
        public string? DetailedReportTemplatePath { get; set; }
        public string? SummaryReportTemplatePath { get; set; }
        public string? DisclaimerReportTemplatePath { get; set; }
    }

    public  class ReportPath
    {
        public string? Path { get; set; }
        public string? SummaryReportPath { get; set; }
        public string? DisclaimerReportPath { get; set; }
        public string? MergedReportPath { get; set; }

    }
    public class Records
    {
        public string? Path { get; set; }
    }
    public class Images
    {
        public string? Downward { get; set; }
        public string? Upward { get; set; }
        public string? DownArrow { get; set; }
        public string? UpArrow { get; set; }
    }
    public class RichTextRender
    {
        public string? DisclaimerRender { get; set; }

        public string? SummaryRender { get; set; }
    }
    public class Passwords
    {
        public string? TemplatePassword { get; set; }

    }
}
