using Azure.Core;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net.Http.Headers;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    internal class Auth
    {
        public GraphServiceClient graphAuth(ILogger log)
        {
            IConfiguration config = new ConfigurationBuilder()
           .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
           .AddEnvironmentVariables()
           .Build();

            log.LogInformation("graphAuth processed a request.");

            var scopes = new string[] { "https://graph.microsoft.com/.default" };
            var keyVaultUrl = config["keyVaultUrl"];
            var clientId = config["clientId"];
            var tenantid = config["tenantid"];
            var secretName = config["secretName"];

            //SecretClientOptions options = new SecretClientOptions()
            //{
            //    Retry =
            //{
            //    Delay= TimeSpan.FromSeconds(2),
            //    MaxDelay = TimeSpan.FromSeconds(16),
            //    MaxRetries = 5,
            //    Mode = RetryMode.Exponential
            // }
            //};
            
            //var client = new SecretClient(new Uri(keyVaultUrl), new DefaultAzureCredential(), options);

            //KeyVaultSecret secret = client.GetSecret(secretName);

            IConfidentialClientApplication confidentialClientApplication = ConfidentialClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantid)
            .WithClientSecret("j448Q~poewWoAa6X8RZaONmf0HmXH_KdaVDDDdzC")
            .Build();

            // Build the Microsoft Graph client. As the authentication provider, set an async lambda
            // which uses the MSAL client to obtain an app-only access token to Microsoft Graph,
            // and inserts this access token in the Authorization header of each API request. 
            GraphServiceClient graphServiceClient =
                new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                {

                    // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                    var authResult = await confidentialClientApplication
                        .AcquireTokenForClient(scopes)
                        .ExecuteAsync();

                    // Add the access token in the Authorization header of the API request.
                    requestMessage.Headers.Authorization =
                        new AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                })
                );
            return graphServiceClient;
        }
    }
}