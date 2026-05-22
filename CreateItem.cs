using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Newtonsoft.Json;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public class CreateItem
    {
        private readonly ILogger<CreateItem> _logger;
        public CreateItem(ILogger<CreateItem> logger)
        {
            _logger = logger;
        }

        [Function("CreateItem")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req)
        {
            _logger.LogInformation("CreateItem processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: true, reloadOnChange: true).AddEnvironmentVariables().Build();

            string connectionString = config["AzureWebJobsStorage"];

            string SecurityCategory = data?.SecurityCategory;
            string SpaceName = data?.SpaceName;
            string SpaceNameFR = data?.SpaceNameFR;
            string Owner1 = data?.Owner1;
            string SpaceDescription = data?.SpaceDescription;
            string SpaceDescriptionFR = data?.SpaceDescriptionFR;
            string TemplateTitle = data?.TemplateTitle;
            string TeamPurpose = data?.TeamPurpose;
            string BusinessJustification = data?.BusinessJustification;
            string RequesterName = data?.RequesterName;
            string RequesterEmail = data?.RequesterEmail;
            string Status = data?.Status;
            //string Members = data?.Members;

            if (TeamPurpose == "ERROR")
            {
                return new BadRequestResult();
            }

            var listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"SecurityCategory", SecurityCategory},
                        {"Title", SpaceName},
                        {"SpaceNameFR", SpaceNameFR},
                        {"Owner1", Owner1},
                        {"SpaceDescription", SpaceDescription},
                        {"SpaceDescriptionFR", SpaceDescriptionFR},
                        {"TemplateTitle", TemplateTitle},
                        {"TeamPurpose", TeamPurpose},
                        {"BusinessJustification", BusinessJustification},
                        {"RequesterName", RequesterName},
                        {"RequesterEmail", RequesterEmail},
                        {"Status", Status}//,
                        //{"Members", Members}
                    }
                }
            };

            string ValidationErrors = Common.ValidateInput(listItem.Fields);

            if (ValidationErrors == "")
            {
                try
                {
                    ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(_logger);
                    GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                    await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items.PostAsync(listItem);

                    // send item to email queue to trigger email to user
                    Common.InsertMessageAsync2(connectionString, "email", listItem, _logger).GetAwaiter().GetResult();

                    return new OkResult();
                }
                catch (Exception e)
                {
                    _logger.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null)
                        _logger.LogError($"InnerException: {e.InnerException.Message}");
                    _logger.LogError($"StackTrace: {e.StackTrace}");

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