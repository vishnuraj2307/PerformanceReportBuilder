using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using PRB.Services;
using RulesEngine.Models;
using Serilog;

namespace PRB.Repository
{
    public interface IRuleExecutor
    {
        Task<string> GetHomeEngine(object data, string workflowName);
       
    }
    public class RuleExecutor : IRuleExecutor
    {
        protected readonly AppSettings appsettings;
        private string connectionStr = string.Empty;

        private List<Workflow> items = new List<Workflow>();
        public RuleExecutor( IConfiguration configuration)
        {
            //appsettings = configuration.GetSection()
           // connectionStr = configuration.GetConnectionString("MyDBConnection");
            var WorkFlowRules = "";
            using(StreamReader r = new StreamReader(Path.Combine(Directory.GetCurrentDirectory(),"RuleEngine.json"))) 
            {
                WorkFlowRules=r.ReadToEnd();
                items = JsonConvert.DeserializeObject<List<Workflow>>(WorkFlowRules);
            }
        }

        public async Task<string> GetHomeEngine(object data, string workflowName)
        {
            dynamic response = null;
            var re = new RulesEngine.RulesEngine(items.ToArray());
            Log.Verbose("Initiating Rule Engine...");
            var resultList = await re.ExecuteAllRulesAsync(workflowName, data);
            Log.Information("RuleExecutor returned the values.\n");
            foreach (var result in resultList)
            {
                if(result.IsSuccess)
                {
                    response = result.Rule.Actions.OnSuccess.Context.GetValueOrDefault("Expression");
                    break;
                }
               
            }
            if (response != null)
            {
                return response;
            }
            else
            {
                return "Error";
            }
          
        }
    }
}
