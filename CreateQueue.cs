using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using Microsoft.Azure.Functions.Worker;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class CreateQueue
    {
        private readonly ILogger<CreateQueue> _logger;
        public CreateQueue(ILogger<CreateQueue> logger)
        {
            _logger = logger;
        }

        [Function("CreateQueue")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req)
        {
            _logger.LogInformation("CreateQueue received a request.");

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

            Common.InsertMessageAsync(connectionString, queueName, JsonConvert.SerializeObject(listItem.Fields.AdditionalData), _logger).GetAwaiter().GetResult();

            _logger.LogInformation("CreateQueue processed a request.");
            return new OkResult();
        }
    }
}