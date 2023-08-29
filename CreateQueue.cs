using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

using Newtonsoft.Json;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public static class CreateQueue
    {
        [FunctionName("CreateQueue")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("CreateQueue received a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string Comment = data?.Comment;
            string connectionString = config["AzureWebJobsStorage"];
            string ItemId = data?.Id;
            string queueName = "list";
            string Status = data?.Status;

            ListItem listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"Id", ItemId},
                        {"Status", Status},
                        {"Comment", Comment}
                    }
                }
            };

            Common.InsertMessageAsync(connectionString, queueName, JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log).GetAwaiter().GetResult();

            log.LogInformation("CreateQueue processed a request.");
            return new OkResult();
        }
    }
}