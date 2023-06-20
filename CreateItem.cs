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
using System.Collections.Generic;
using Microsoft.Extensions.Configuration;
using static appsvc_fnc_dev_scw_list_dotnet001.Auth;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public static class CreateItem
    {
        [FunctionName("CreateItem")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = null)] HttpRequest req, ILogger log)
        {
            log.LogInformation("CreateItem processed a request.");

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
                    ROPCConfidentialTokenCredential auth = new ROPCConfidentialTokenCredential(log);
                    GraphServiceClient graphAPIAuth = new GraphServiceClient(auth);

                    await graphAPIAuth.Sites[config["siteId"]].Lists[config["listId"]].Items.Request().AddAsync(listItem);

                    // send item to email queue to trigger email to user
                    Common.InsertMessageAsync(connectionString, "email", JsonConvert.SerializeObject(listItem.Fields.AdditionalData), log).GetAwaiter().GetResult();

                    return new OkResult();
                }
                catch (Exception e)
                {
                    log.LogError($"Message: {e.Message}");
                    if (e.InnerException is not null)
                        log.LogError($"InnerException: {e.InnerException.Message}");
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