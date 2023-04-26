using System;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class UpdateItem
    {
        [FunctionName("UpdateItem")]
        public async Task RunAsync([QueueTrigger("list", Connection = "AzureWebJobsStorage")] string myQueueItem, ILogger log)
        {
            log.LogInformation("_UpdateItem received a request.");

            dynamic data = JsonConvert.DeserializeObject(myQueueItem);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string connectionString = config["AzureWebJobsStorage"];

            string ItemId = data?.Id;
            string Status = data?.Status;
            string ApprovedDate = DateTime.Now.ToLocalTime().ToString();
            string Comment = data?.Comment;

            var fieldValueSet = new FieldValueSet
            {
                AdditionalData = new Dictionary<string, object>()
                    {
                        {"Status", Status},
                        {"ApprovedDate", ApprovedDate},
                        {"Comment", Comment}
                    }
            };

            string ValidationErrors = Common.ValidateInput(fieldValueSet);

            if (ValidationErrors == "")
            {
                try
                {
                    ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
                    GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                    await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items[ItemId].Fields.Request().UpdateAsync(fieldValueSet);

                    string queueName = string.Empty;
                    if (Status == "Approved")
                        queueName = "sitecreation";
                    else if (Status == "Rejected")
                        queueName = "email";

                    log.LogInformation($"queueName: {queueName}");

                    if (queueName != string.Empty)
                    {
                        ListItem listItem = await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items[ItemId].Request().GetAsync();
                        Common.InsertMessageAsync2(connectionString, queueName, listItem, log).GetAwaiter().GetResult();
                    }
                }
                catch (Exception e)
                {
                    log.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null)
                        log.LogError($"InnerException: {e.InnerException.Message}");
                }
            }
            else
            {
                throw new Exception($"ValidationErrors: {ValidationErrors}");
            }
        }
    }
}
