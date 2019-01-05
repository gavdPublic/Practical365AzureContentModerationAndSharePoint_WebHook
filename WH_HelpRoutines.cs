using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using System;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security;

namespace WHFunction
{
    class WHHelpRoutines
    {
        public static string GetValidationToken(HttpRequestMessage req, TraceWriter log)
        {
            log.Info($"---- Registration: Validating Token");
            string strReturn = string.Empty;

            try
            {
                strReturn = req.GetQueryNameValuePairs()
                    .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                    .Value;
                log.Info($"---- Registration: Validation Token Found : " + strReturn);
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetValidationToken: " + ex.ToString());
            }

            return strReturn;
        }

        public static HttpResponseMessage RegisterWebHook(HttpRequestMessage req, string validationToken, TraceWriter log)
        {
            log.Info($"---- Registration: Registering WebHook");
            HttpResponseMessage resReturn = null;

            try
            {
                resReturn = req.CreateResponse(HttpStatusCode.OK);
                resReturn.Content = new StringContent(validationToken);
                log.Info($"---- Registration: WebHook Registration succeeded");
            }
            catch (Exception ex)
            {
                log.Info($"Error in RegisterWebHook: " + ex.ToString());
            }

            return resReturn;
        }

        public static ClientContext LoginSharePoint(string BaseUrl, TraceWriter log)
        {
            // Login using UserOnly Credentials (User Name and User PW)
            ClientContext cntReturn = null;

            try
            {
                string myUserName = ConfigurationManager.AppSettings["whSpUserName"];
                string myPassword = ConfigurationManager.AppSettings["whSpUserPw"];

                SecureString securePassword = new SecureString();
                foreach (char oneChar in myPassword) securePassword.AppendChar(oneChar);
                SharePointOnlineCredentials myCredentials = new SharePointOnlineCredentials(myUserName, securePassword);

                cntReturn = new ClientContext(BaseUrl);
                cntReturn.Credentials = myCredentials;
            }
            catch (Exception ex)
            {
                log.Info($"Error in LoginSharePoint: " + ex.ToString());
            }

            return cntReturn;
        }

        //public static ClientContext LoginSharePoint(BDObjects.IntranetLanguages UrlLanguage, string BaseUrl)
        //{
            // Login using AppOnly Credentials (ClientID, ClientSecret and Realm)
            //ClientContext cntReturn = null;

            //string myClientId = string.Empty;
            //string myClientSecret = string.Empty;
            //if (UrlLanguage == BDObjects.IntranetLanguages.Engels)
            //{
            //    myClientId = ConfigurationManager.AppSettings["FCAuthorizationClientID_EN"]; ;
            //    myClientSecret = ConfigurationManager.AppSettings["FCAuthorizationSecret_EN"];
            //}
            //else if (UrlLanguage == BDObjects.IntranetLanguages.Netherlands)
            //{
            //    myClientId = ConfigurationManager.AppSettings["FCAuthorizationClientID_NL"]; ;
            //    myClientSecret = ConfigurationManager.AppSettings["FCAuthorizationSecret_NL"];
            //}
            //string myRealm = ConfigurationManager.AppSettings["FCAuthorizationRealm"];

            //if (string.IsNullOrEmpty(myClientId) == false &&
            //    string.IsNullOrEmpty(myClientSecret) == false &&
            //    string.IsNullOrEmpty(myRealm) == false)
            //{
            //    OfficeDevPnP.Core.AuthenticationManager myAuthManager = new OfficeDevPnP.Core.AuthenticationManager();
            //    cntReturn = myAuthManager.GetAppOnlyAuthenticatedContext(BaseUrl, myRealm, myClientId, myClientSecret);
            //}

        //    return cntReturn;
        //}

        public static ChangeQuery GetChangeQueryNew(string ListId, TraceWriter log)
        {
            // ATENTION: Change Token made for only the last minute and for new Items !!

            ChangeToken lastChangeToken = new ChangeToken();
            ChangeQuery myChangeQuery = new ChangeQuery(false, false);

            try
            {
            lastChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.AddMinutes(-1).ToUniversalTime().Ticks.ToString());
            ChangeToken newChangeToken = new ChangeToken();
            newChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.ToUniversalTime().Ticks.ToString());
            myChangeQuery.Item = true;  // Get only Item changes
            myChangeQuery.Add = true;   // Get only the new Items
            myChangeQuery.ChangeTokenStart = lastChangeToken;
            myChangeQuery.ChangeTokenEnd = newChangeToken;
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetChangeQueryNew: " + ex.ToString());
            }

            return myChangeQuery;
        }

        public static ChangeQuery GetChangeQueryNewUpdate(string ListId, TraceWriter log)
        {
            // ATENTION: Change Token made for only the last minute and for new Items !!

            ChangeToken lastChangeToken = new ChangeToken();
            ChangeQuery myChangeQuery = new ChangeQuery(false, false);

            try
            {
                lastChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.AddMinutes(-1).ToUniversalTime().Ticks.ToString());
                ChangeToken newChangeToken = new ChangeToken();
                newChangeToken.StringValue = string.Format("1;3;{0};{1};-1", ListId, DateTime.Now.ToUniversalTime().Ticks.ToString());
                myChangeQuery.Item = true;  // Get only Item changes
                myChangeQuery.Add = true;   // Get only new and updated Items
                myChangeQuery.Update = true;   // Get only new and updated Items
                myChangeQuery.ChangeTokenStart = lastChangeToken;
                myChangeQuery.ChangeTokenEnd = newChangeToken;
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetChangeQueryNewUpdate: " + ex.ToString());
            }

            return myChangeQuery;
        }

        public static List GetListWebHook(ClientContext SPClientContext, TraceWriter log)
        {
            List lstReturn = null;

            try
            {
                string listTitle = ConfigurationManager.AppSettings["whSpListTitle"];
                Web spWeb = SPClientContext.Site.RootWeb;
                lstReturn = spWeb.Lists.GetByTitle(listTitle);
                SPClientContext.Load(lstReturn);
                SPClientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetListWebHook: " + ex.ToString());
            }

            return lstReturn;
        }

        public static ChangeCollection GetAllChanges(List ListWebHook, ChangeQuery ListChangeQuery, 
                                                    ClientContext SPClientContext, TraceWriter log)
        {
            ChangeCollection chcReturn = null;

            try
            {
                chcReturn = ListWebHook.GetChanges(ListChangeQuery);
                SPClientContext.Load(chcReturn);
                SPClientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetAllChanges: " + ex.ToString());
            }

            return chcReturn;
        }

        public static ListItem GetItemChanged(List ListWebHook, Change GivenChange, 
                                                    ClientContext SPClientContext, TraceWriter log)
        {
            ListItem litReturn = null;
                
            try
            {
                litReturn = ListWebHook.GetItemById((GivenChange as ChangeItem).ItemId);
                SPClientContext.Load(litReturn);
                SPClientContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                log.Info($"Error in GetItemChanged: " + ex.ToString());
            }

            return litReturn;
        }
    }
}
