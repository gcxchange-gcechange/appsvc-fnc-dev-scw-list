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
using Microsoft.IdentityModel.Tokens;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public static class CreateItem
    {
        [FunctionName("CreateItem")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
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
                        {"Status", Status}
                    }
                }
            };

            string ValidationErrors = ValidateInput(listItem);

            if (ValidationErrors == "")
            {
                Auth auth = new();
                GraphServiceClient graphAPIAuth = auth.graphAuth(log);
                await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items.Request().AddAsync(listItem);
                return new OkResult();
            }
            else
            {
                return new BadRequestObjectResult(ValidationErrors);
            }
        }

        /// <summary>
        /// Ensure that all fields have a non-empty value
        /// </summary>
        /// <param name="listItem"></param>
        /// <returns>String list of validation errors, otherwise empty string</returns>
        private static string ValidateInput(ListItem listItem)
        {
            string ValidationErrors = String.Empty;
            foreach (string k in listItem.Fields.AdditionalData.Keys)
            {
                if ((listItem.Fields.AdditionalData[k] is null) || string.IsNullOrEmpty(listItem.Fields.AdditionalData[k].ToString().Trim()))
                {
                    ValidationErrors += string.Format("Field {0} cannot be blank.", k) + Environment.NewLine;
                }
            }

            return ValidationErrors;
        }
    }
}