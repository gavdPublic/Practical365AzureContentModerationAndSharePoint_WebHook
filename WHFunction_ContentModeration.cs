using Microsoft.Azure.CognitiveServices.ContentModerator;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using WHFunction;

namespace SPWebHook_AzureFunction_ContentModeration
{
    public static class WHFunction_ContentModeration
    {
        [FunctionName("WHFunction_ContentModeration")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, 
                                                            "post", Route = null)]HttpRequestMessage req, 
                                                            TraceWriter log)
        {
            log.Info($"WebHook Function called at: { DateTime.Now }");

            // WebHook Registration
            string validationToken = WHHelpRoutines.GetValidationToken(req, log);
            if (validationToken != null)
                return WHHelpRoutines.RegisterWebHook(req, validationToken, log);
            else
                log.Info($"---- Registration: Registering WebHook already done");

            // Changes
            var myContent = await req.Content.ReadAsStringAsync();
            var allNotifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(myContent).Value;

            if (allNotifications.Count > 0)
            {
                foreach (var oneNotification in allNotifications)
                {
                    // Login in SharePoint
                    string baseUrl = ConfigurationManager.AppSettings["whSpBaseUrl"];
                    ClientContext SPClientContext = WHHelpRoutines.LoginSharePoint(baseUrl, log);

                    // Get the Changes
                    GetChanges(SPClientContext, oneNotification.Resource, log);
                }
            }

            return new HttpResponseMessage(HttpStatusCode.OK);
        }

        static void GetChanges(ClientContext SPClientContext, string ListId, TraceWriter log)
        {
            // Get the List where the WebHook is working
            List myListWebHook = WHHelpRoutines.GetListWebHook(SPClientContext, log);

            // Get the Change Query
            ChangeQuery myChangeQuery = WHHelpRoutines.GetChangeQueryNew(ListId, log); // Only new items and for the last one minute

            // Get all the Changes
            ChangeCollection allChanges = WHHelpRoutines.GetAllChanges(myListWebHook, myChangeQuery, SPClientContext, log);

            foreach (Change oneChange in allChanges)
            {
                if (oneChange is ChangeItem)
                {
                    // Get what is changed
                    ListItem itemChanged = WHHelpRoutines.GetItemChanged(myListWebHook, oneChange, SPClientContext, log);
                    log.Info($"itemChangedID - " + itemChanged.Id.ToString());

                    // Do something with the Item Changed
                    DoSomething(SPClientContext, itemChanged, log);
                }
            }
        }

        static void DoSomething(ClientContext SPClientContext, ListItem ItemChanged, TraceWriter log)
        {
            string disBody = string.Empty;

            // Read the post Body
            disBody = ItemChanged["Body"].ToString();
            log.Info($"---- Body: " + disBody);

            // Make a stream of the body text
            byte[] disBodyByteArray = System.Text.Encoding.UTF8.GetBytes(disBody);
            MemoryStream disBodyStream = new MemoryStream(disBodyByteArray);
            log.Info($"---- Body For Moderation: " + disBody);

            // Get the Moderation from Azure and modify the Body
            AzureContentModerationResults myModeration = CheckContentForModeration(disBodyStream);
            string myModifiedBody = ModifyText(myModeration.AutoCorrectedText, myModeration);
            log.Info($"---- Body Moderated: " + myModifiedBody);

            // Send back the modified Body to SharePoint
            ItemChanged["Body"] = myModifiedBody;
            ItemChanged.Update();
            SPClientContext.ExecuteQuery();
            log.Info($"---- Item Moderated");
        }

        static AzureContentModerationResults CheckContentForModeration(MemoryStream TextToModerate)
        {
            string AzureContModBaseURL = "https://westeurope.api.cognitive.microsoft.com";
            string ContModSubscriptionKey = "dabb15da19a14baea06a71bfxxxxxxxx";

            ContentModeratorClient contModClient = new ContentModeratorClient(new ApiKeyServiceClientCredentials(ContModSubscriptionKey));
            contModClient.Endpoint = AzureContModBaseURL;
            var objResult = contModClient.TextModeration.ScreenText("text/plain", TextToModerate, "eng", true, true, null, true);
            string jsonResult = JsonConvert.SerializeObject(objResult, Formatting.Indented);

            return JsonConvert.DeserializeObject<AzureContentModerationResults>(jsonResult);
        }

        static string ModifyText(string InputString, AzureContentModerationResults ModerationResult)
        {
            foreach (Term oneWord in ModerationResult.Terms)
            {
                string termModified = oneWord.term.Remove(1);
                for (int charCount = 1; charCount < oneWord.term.Length; charCount++)
                {
                    termModified += "*";
                }
                InputString = InputString.Replace(oneWord.term, termModified);
            }

            return InputString;
        }
    }

    public class AzureContentModerationResults
    {
        public string OriginalText { get; set; }
        public string NormalizedText { get; set; }
        public string AutoCorrectedText { get; set; }
        public object Misrepresentation { get; set; }
        public Classification Classification { get; set; }
        public Status Status { get; set; }
        public PII PII { get; set; }
        public string Language { get; set; }
        public Term[] Terms { get; set; }
        public string TrackingId { get; set; }
    }

    public class Classification
    {
        public Category1 Category1 { get; set; }
        public Category2 Category2 { get; set; }
        public Category3 Category3 { get; set; }
        public bool ReviewRecommended { get; set; }
    }

    public class Category1
    {
        public float Score { get; set; }
    }

    public class Category2
    {
        public float Score { get; set; }
    }

    public class Category3
    {
        public float Score { get; set; }
    }

    public class Status
    {
        public int Code { get; set; }
        public string Description { get; set; }
        public object Exception { get; set; }
    }

    public class PII
    {
        public Email[] Email { get; set; }
        public SSN[] SSN { get; set; }
        public IPA[] IPA { get; set; }
        public Phone[] Phone { get; set; }
        public Address[] Address { get; set; }
    }

    public class Email
    {
        public string Detected { get; set; }
        public string SubType { get; set; }
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class SSN
    {
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class IPA
    {
        public string SubType { get; set; }
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class Phone
    {
        public string CountryCode { get; set; }
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class Address
    {
        public string Text { get; set; }
        public int Index { get; set; }
    }

    public class Term
    {
        public int Index { get; set; }
        public int OriginalIndex { get; set; }
        public int ListId { get; set; }
        public string term { get; set; }
    }
}
