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

            IConfiguration config = new ConfigurationBuilder()
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

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

            var listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
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
                        {"Status", Status},
                    }
                }
            };

            string ValidationErrors = Validation.ValidateInput(listItem.Fields);

            if (ValidationErrors == "")
            {
                try
                {
                    Auth auth = new();
                    GraphServiceClient graphAPIAuth = auth.graphAuth(log);
                    await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items.Request().AddAsync(listItem);
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