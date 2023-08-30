using System;
using System.Collections.Generic;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class UpdateStatus
    {
        [FunctionName("UpdateStatus")]
        public async Task RunAsync([QueueTrigger("status", Connection = "AzureWebJobsStorage")]string myQueueItem, ILogger log)
        {
            log.LogInformation("UpdateStatus received a request.");

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string siteId = config["siteId"];
            string listId = config["listId"];

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            string ItemId = data?.Id;
            string Status = data?.Status;

            var fieldValueSet = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>()
                    {
                        {"Status", Status}
                    }
            };

            try
            {
                ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
                GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                await graphAPIAuth.Sites[siteId].Lists[listId].Items[ItemId].Fields.Request().UpdateAsync(fieldValueSet);
            }
            catch (Exception e)
            {
                log.LogError($"Message: {e.Message}");
                if (e.InnerException is not null)
                    log.LogError($"InnerException: {e.InnerException.Message}");
                log.LogError($"StackTrace: {e.StackTrace}");
            }

            log.LogInformation("UpdateStatus processed a request.");
        }
    }
}