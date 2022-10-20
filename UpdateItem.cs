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
    public class SpaceRequest
    {
        public string SpaceName { get; set; }
        public string SpaceNameFR { get; set; }
        public string Owner1 { get; set; }
        public string SpaceDescription { get; set; }
        public string SpaceDescriptionFR { get; set; }
        public string TemplateTitle { get; set; }
        public string TeamPurpose { get; set; }
        public string BusinessJustification { get; set; }
        public string RequesterName { get; set; }
        public string RequesterEmail { get; set; }
        public string Status { get; set; }
        public string ApprovedDate { get; set; }
        public string Comment { get; set; }
   }

    public static class UpdateItem
    {
        [FunctionName("UpdateItem")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("UpdateItem processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            IConfiguration config = new ConfigurationBuilder()
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

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
                        queueName = "site-creation";
                    else if (Status == "Rejected")
                        queueName = "email";

                    if (queueName != string.Empty)
                    {
                        //send item to queue
                        var connectionString = config["AzureWebJobsStorage"];
                        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
                        CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                        CloudQueue queue = queueClient.GetQueueReference(queueName);

                        ListItem listItem = await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items[ItemId].Request().GetAsync();
                        InsertMessageAsync(listItem, queue, log).GetAwaiter().GetResult();

                        log.LogInformation("Sent request to queue successful.");
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

        private static async Task InsertMessageAsync(ListItem listItem, CloudQueue theQueue, ILogger log)
        {
            SpaceRequest request = new SpaceRequest();

            request.SpaceName = listItem.Fields.AdditionalData["Title"].ToString();
            request.SpaceNameFR = listItem.Fields.AdditionalData["SpaceNameFR"].ToString();
            request.Owner1 = listItem.Fields.AdditionalData["Owner1"].ToString();
            request.SpaceDescription = listItem.Fields.AdditionalData["SpaceDescription"].ToString();
            request.SpaceDescriptionFR = listItem.Fields.AdditionalData["SpaceDescriptionFR"].ToString();
            request.TemplateTitle = listItem.Fields.AdditionalData["TemplateTitle"].ToString();
            request.TeamPurpose = listItem.Fields.AdditionalData["TeamPurpose"].ToString();
            request.BusinessJustification = listItem.Fields.AdditionalData["BusinessJustification"].ToString();
            request.RequesterName = listItem.Fields.AdditionalData["RequesterName"].ToString();
            request.RequesterEmail = listItem.Fields.AdditionalData["RequesterEmail"].ToString();
            request.Status = listItem.Fields.AdditionalData["Status"].ToString();
            request.ApprovedDate = listItem.Fields.AdditionalData["ApprovedDate"].ToString();
            request.Comment = listItem.Fields.AdditionalData["Comment"].ToString();

            string serializedMessage = JsonConvert.SerializeObject(request);

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await theQueue.AddMessageAsync(message);
        }
    }
}