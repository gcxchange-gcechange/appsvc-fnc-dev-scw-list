using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;
using Microsoft.WindowsAzure.Storage.Queue;
using Microsoft.WindowsAzure.Storage;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    

    public static class UpdateItem
    {
        [FunctionName("UpdateItem")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("UpdateItem processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string connectionString = config["AzureWebJobsStorage"];

            string ItemId = data?.ItemId;
            string Status = data?.Status;
            string ApprovedDate = data?.ApprovedDate ?? DateTime.Now.ToLocalTime().ToString();
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

            string ValidationErrors = Validation.ValidateInput(fieldValueSet);

            if (ValidationErrors == "")
            {
               try
                {
                    Auth auth = new();
                    GraphServiceClient graphAPIAuth = auth.graphAuth(log);
                    await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items[ItemId].Fields.Request().UpdateAsync(fieldValueSet);

                    string queueName = string.Empty;
                    if (Status == "Approved")
                        queueName = "sitecreation";
                    else if (Status == "Rejected")
                        queueName = "email";

                    if (queueName != string.Empty)
                    {
                        // send item to queue
                        ListItem listItem = await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items[ItemId].Request().GetAsync();
                        Common.InsertMessageAsync(connectionString, queueName, listItem, log).GetAwaiter().GetResult();
                    }

                    return new OkResult();
                }
                catch (Exception e)
                {
                    log.LogInformation(e.Message);
                    if (e.InnerException is not null)
                        log.LogInformation(e.InnerException.Message);
                    return new BadRequestResult();
                }
            }
            else
            {
                return new BadRequestObjectResult(ValidationErrors);
            }
        }

       
    }
}