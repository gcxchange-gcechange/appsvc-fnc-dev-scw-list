using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;
using Microsoft.Azure.Functions.Worker;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class UpdateItem
    {
        private readonly ILogger<UpdateItem> _logger;
        public UpdateItem(ILogger<UpdateItem> logger)
        {
            _logger = logger;
        }

        [Function("UpdateItem")]
        public async Task RunAsync([QueueTrigger("list", Connection = "AzureWebJobsStorage")] string myQueueItem)
        {
            _logger.LogInformation("UpdateItem received a request.");

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
                    ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(_logger);
                    GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                    await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items[ItemId].Fields.PatchAsync(fieldValueSet);

                    string queueName = string.Empty;
                    if (Status == "Approved")
                        queueName = "sitecreation";
                    else if (Status == "Rejected")
                        queueName = "email";

                    _logger.LogInformation($"queueName: {queueName}");

                    if (queueName != string.Empty)
                    {
                        ListItem listItem = await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items[ItemId].GetAsync();
                        Common.InsertMessageAsync2(connectionString, queueName, listItem, _logger).GetAwaiter().GetResult();
                    }
                }
                catch (Exception e)
                {
                    _logger.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null)
                        _logger.LogError($"InnerException: {e.InnerException.Message}");
                    _logger.LogError($"StackTrace: {e.StackTrace}");
                }
            }
            else
            {
                throw new Exception($"ValidationErrors: {ValidationErrors}");
            }


            _logger.LogInformation("UpdateItem processed a request.");
        }
    }
}
