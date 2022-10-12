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
                    log.LogInformation("123");
                    await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items[ItemId].Fields.Request().UpdateAsync(fieldValueSet);
                    log.LogInformation("456");
                    return new OkResult();
                }
                catch (Exception e)
                {
                    log.LogInformation(e.Message);
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