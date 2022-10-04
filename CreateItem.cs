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
using System.Net.NetworkInformation;
using Microsoft.Extensions.Configuration;

// Inside of it, you need one function call : createitem that need to use the graph api to create an item in a sharepoint list.
// To get this working, you will need to create all resource in Azure, create a team site call App-SCW2 and a sharepoint list call: Request.
// ALso, you will need to create an app registration with the api permission so the app can write to the sp list.

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    public static class CreateItem
    {
        [FunctionName("CreateItem")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("CreateItem processed a request.");

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);

            IConfiguration config = new ConfigurationBuilder()
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            var listItem = new ListItem
            {
                Fields = new FieldValueSet
                {
                    AdditionalData = new Dictionary<string, object>()
                    {
                        {"Title", "Space Name"},
                        {"SpaceNameFR", "Space Name FR"},
                        {"Owner1", "Owner 1"},
                        {"SpaceDescription", "Space Description"},
                        {"TemplateTitle", "Template Title"},
                        {"SpaceDescriptionFR", "Space Description FR"},
                        {"TeamPurpose", "Team Purpose"},
                        {"BusinessJustification", "Business Justification"},
                        {"RequesterName", "Requester Name"},
                        {"RequesterEmail", "Requester Email"},
                        {"Status", "Status"}
                    }
                }
            };

            Auth auth = new();
            GraphServiceClient graphAPIAuth = auth.graphAuth(log);

            await graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Items
            .Request()
            .AddAsync(listItem);

            //var listColumns = graphAPIAuth.Sites[config["SiteId"]].Lists[config["ListId"]].Columns.Request().GetAsync();
            //return new OkObjectResult(JsonConvert.SerializeObject(listColumns, Formatting.None, new JsonSerializerSettings() { PreserveReferencesHandling = PreserveReferencesHandling.Objects }).ToString());

            return new OkObjectResult("Huzzah!");
        }
    }
}