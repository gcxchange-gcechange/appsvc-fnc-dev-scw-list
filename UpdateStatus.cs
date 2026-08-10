using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;
using Microsoft.Azure.Functions.Worker;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class UpdateStatus
    {
        private readonly ILogger<UpdateStatus> _logger;
        public UpdateStatus(ILogger<UpdateStatus> logger)
        {
            _logger = logger;
        }

        [Function("UpdateStatus")]
        public async Task RunAsync([QueueTrigger("status", Connection = "AzureWebJobsStorage")]string myQueueItem)
        {
            _logger.LogInformation("UpdateStatus received a request.");

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
                ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(_logger);
                GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                await graphAPIAuth.Sites[siteId].Lists[listId].Items[ItemId].Fields.PatchAsync(fieldValueSet);
            }
            catch (Exception e)
            {
                _logger.LogError($"Message: {e.Message}");
                if (e.InnerException is not null)
                    _logger.LogError($"InnerException: {e.InnerException.Message}");
                _logger.LogError($"StackTrace: {e.StackTrace}");
            }

            _logger.LogInformation("UpdateStatus processed a request.");
        }
    }
}