using System;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Web.Script.Serialization;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;

namespace GxDcCPSEmailNotificationfnc
{
    public static class EmailQueue
    {
        [FunctionName("EmailQueue")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            // parse query parameter
            string emails = "";
            string comments = "";
            string status = "";
            string displayName = "";
            string requesterName = "";
            string requesterEmail = "";

            // Get request body
            dynamic data = await req.Content.ReadAsAsync<object>();
            displayName = data?.name;
            emails = data?.emails;
            comments = data?.comments;
            status = data?.status;
            requesterName = data?.requesterName;
            requesterEmail = data?.requesterEmail;

            //send message to queue
            var connectionString = ConfigurationManager.AppSettings["AzureWebJobsStorage"];

            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(connectionString);
            CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
            CloudQueue queue = queueClient.GetQueueReference("email-info");
            log.Info($"site name {displayName}");
            InsertMessageAsync(queue, displayName, emails, status, comments, requesterEmail, log).GetAwaiter().GetResult();
            log.Info($"Sent request to queue successful.");

            return displayName == null
                ? req.CreateResponse(HttpStatusCode.BadRequest, "Please pass a name on the query string or in the request body")
                : req.CreateResponse(HttpStatusCode.OK, "Hello " + displayName);
        }

        public static async Task InsertMessageAsync(CloudQueue theQueue, string dispalyName, string emails, string status, string comments, string requesterEmail, TraceWriter log)
        {
            SiteInfo siteInfo = new SiteInfo();

            switch (status)
            {
                case "Submitted":
                    siteInfo.displayName = dispalyName;
                    siteInfo.status = status;
                    siteInfo.requesterEmail = requesterEmail;
                    log.Info($"display is {dispalyName}, status is {status}");
                    break;

                case "Rejected":
                    siteInfo.displayName = dispalyName;
                    siteInfo.status = status;
                    siteInfo.comments = comments;
                    siteInfo.requesterEmail = requesterEmail;
                    break;

                case "Notif_HD":
                    siteInfo.displayName = dispalyName;
                    siteInfo.status = status;
                    break;
            }

            string serializedMessage = JsonConvert.SerializeObject(siteInfo);
            if (await theQueue.CreateIfNotExistsAsync())
            {
                log.Info("The queue was created.");
            }

            CloudQueueMessage message = new CloudQueueMessage(serializedMessage);
            await theQueue.AddMessageAsync(message);
        }
    }
}
