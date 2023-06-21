using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Newtonsoft.Json;
using System.Threading.Tasks;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using System;

namespace appsvc_fnc_dev_scw_list_dotnet001
{
    internal class Common
    {

        public class SpaceRequest
        {
            public string Id { get; set; }
            public string SecurityCategory { get; set; }
            public string SpaceName { get; set; }
            public string SpaceNameFR { get; set; }
            public string Owner1 { get; set; }
            public string SpaceDescription { get; set; }
            public string SpaceDescriptionFR { get; set; }
            public string TemplateTitle { get; set; }
            public string TeamPurpose { get; set; }
            public string BusinessJustification { get; set; }
            public string RequesterName { get; set; }
            public string RequesterEmail { get; set; }
            public string Status { get; set; }
            public string ApprovedDate { get; set; }
            public string Comment { get; set; }
            //public string Members { get; set; }
        }

        public static async Task InsertMessageAsync(string connectionString, string queueName, string serializedMessage, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            log.LogInformation($"serializedMessage: {serializedMessage}");

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference(queueName);
            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await queue.AddMessageAsync(message);

            log.LogInformation("InsertMessageAsync processed a request.");
        }

        public static async Task InsertMessageAsync2(string connectionString, string queueName, ListItem listItem, ILogger log)
        {
            log.LogInformation("InsertMessageAsync received a request.");

            // temp, to be refactored

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference(queueName);

            SpaceRequest request = new SpaceRequest();

            request.Id = listItem.Id;
            request.SecurityCategory = listItem.Fields.AdditionalData.Keys.Contains("SecurityCategory") ? listItem.Fields.AdditionalData["SecurityCategory"].ToString() : string.Empty;
            request.SpaceName = listItem.Fields.AdditionalData.Keys.Contains("Title") ? listItem.Fields.AdditionalData["Title"].ToString() : string.Empty;
            request.SpaceNameFR = listItem.Fields.AdditionalData.Keys.Contains("SpaceNameFR") ? listItem.Fields.AdditionalData["SpaceNameFR"].ToString() : string.Empty;
            request.Owner1 = listItem.Fields.AdditionalData.Keys.Contains("Owner1") ? listItem.Fields.AdditionalData["Owner1"].ToString() : string.Empty;
            request.SpaceDescription = listItem.Fields.AdditionalData.Keys.Contains("SpaceDescription") ? listItem.Fields.AdditionalData["SpaceDescription"].ToString() : string.Empty;
            request.SpaceDescriptionFR = listItem.Fields.AdditionalData.Keys.Contains("SpaceDescriptionFR") ? listItem.Fields.AdditionalData["SpaceDescriptionFR"].ToString() : string.Empty;
            request.TemplateTitle = listItem.Fields.AdditionalData.Keys.Contains("TemplateTitle") ? listItem.Fields.AdditionalData["TemplateTitle"].ToString() : string.Empty;
            request.TeamPurpose = listItem.Fields.AdditionalData.Keys.Contains("TeamPurpose") ? listItem.Fields.AdditionalData["TeamPurpose"].ToString() : string.Empty;
            request.BusinessJustification = listItem.Fields.AdditionalData.Keys.Contains("BusinessJustification") ? listItem.Fields.AdditionalData["BusinessJustification"].ToString() : string.Empty;
            request.RequesterName = listItem.Fields.AdditionalData.Keys.Contains("RequesterName") ? listItem.Fields.AdditionalData["RequesterName"].ToString() : string.Empty;
            request.RequesterEmail = listItem.Fields.AdditionalData.Keys.Contains("RequesterEmail") ? listItem.Fields.AdditionalData["RequesterEmail"].ToString() : string.Empty;
            request.Status = listItem.Fields.AdditionalData.Keys.Contains("Status") ? listItem.Fields.AdditionalData["Status"].ToString() : string.Empty;
            request.ApprovedDate = listItem.Fields.AdditionalData.Keys.Contains("ApprovedDate") ? listItem.Fields.AdditionalData["ApprovedDate"].ToString() : string.Empty;
            request.Comment = listItem.Fields.AdditionalData.Keys.Contains("Comment") ? listItem.Fields.AdditionalData["Comment"].ToString() : string.Empty;
            //request.Members = listItem.Fields.AdditionalData.Keys.Contains("Members") ? listItem.Fields.AdditionalData["Members"].ToString() : string.Empty;
            
            string serializedMessage = JsonConvert.SerializeObject(request);

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await queue.AddMessageAsync(message);

            log.LogInformation("InsertMessageAsync processed a request.");
        }


        /// <summary>
        /// Ensure that all fields have a non-empty value
        /// </summary>
        /// <param name="listItem"></param>
        /// <returns>String list of validation errors, otherwise empty string</returns>
        public static string ValidateInput(FieldValueSet fieldValueSet)
        {
            string ValidationErrors = String.Empty;

            foreach (string k in fieldValueSet.AdditionalData.Keys)
            {
                // do not validate the comment field; maybe all validation should be done in the UI anyway??
                if (k.ToLower() != "comment")
                {
                    if ((fieldValueSet.AdditionalData[k] is null) || string.IsNullOrEmpty(fieldValueSet.AdditionalData[k].ToString().Trim()))
                    {
                        ValidationErrors += string.Format("Field {0} cannot be blank.", k) + Environment.NewLine;
                    }
                }
            }

            return ValidationErrors;
        }

    }
}
