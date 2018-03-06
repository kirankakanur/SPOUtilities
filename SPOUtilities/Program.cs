using System;
using System.Collections.Generic;
using System.Text;
using System.Security;
using System.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Taxonomy;
using Microsoft.Office.Server.Search.Administration;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client.Publishing.Navigation;
using Microsoft.SharePoint.Client.Publishing;
using System.Linq;

namespace SPOUtilities
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = ConfigurationManager.AppSettings["SPOSiteUrl"];
            string userName = ConfigurationManager.AppSettings["SPOUserName"];
            string userPassword = ConfigurationManager.AppSettings["SPOPassword"];
            string csvFilePath = ConfigurationManager.AppSettings["CSVFilePath"];
            DisplayJobRunOptions(siteUrl, userName, userPassword, csvFilePath);

            Console.WriteLine("Done processing. Click Enter to close this window.");
            Console.ReadLine();
        }

        public static void DisplayJobRunOptions(string siteUrl, string userName, string password, string csvFilePath)
        {
            string options = "";
            string RunMode = "";
            Console.WriteLine("=== SharePoint Online Utilities ===");
            Console.WriteLine();
            Console.WriteLine("What would you like to do?");
            try
            {
                if (string.IsNullOrEmpty(RunMode))
                {
                    Console.WriteLine("1 - Enable Major and Minor Versions in Document Library");
                    Console.WriteLine("2 - Delete Old Document Versions in Document Library");
                    Console.WriteLine("3 - Get Last Modified Information from Site and Document Library");
                    Console.WriteLine("4 - Add Property Bag values to a List and retrieve them");
                    Console.WriteLine("5 - Update and Retrieve User Profile Properties");
                    Console.WriteLine("6 - Create Terms in Term Store");
                    Console.WriteLine("7 - Work with Managed Properties in Search schema");
                    Console.WriteLine("8 - Get Friendly URLs");
                    Console.WriteLine("");
                    Console.WriteLine("Enter a number from the above list and then click Enter:");
                    RunMode = Console.ReadLine();

                    options = RunMode.GetRightHalf("-").Trim();
                    RunMode = RunMode.GetLeftHalf("-").Trim();
                }

                if (!String.IsNullOrEmpty(RunMode))
                {
                    switch (RunMode)
                    {
                        case "1":
                            EnableVersioningInDocLib(siteUrl, userName, password);
                            break;

                        case "2":
                            DeleteOldDocumentVersions(siteUrl, userName, password, csvFilePath);
                            break;

                        case "3":
                            GetLastModifiedInfo(siteUrl, userName, password, csvFilePath);
                            break;

                        case "4":
                            AddListPropertyBag(siteUrl, userName, password);
                            break;

                        case "5":
                            GetUserProfileProperties(siteUrl, userName, password);
                            break;

                        case "6":
                            CreateTermsInTermStore(siteUrl, userName, password);
                            break;

                        case "7":
                            GetFriendlyUrl(siteUrl, userName, password);
                            break;
                    }
                }
                else
                {
                    Console.WriteLine("Please select an option and try again.");
                    Console.ReadLine();
                }
            }

            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }

            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        private static void GetFriendlyUrl(string siteUrl, string userName, string password)
        {

            //list for saving the urls
            List<string> retVal = new List<string>();

            // get client context
            ClientContext context = GetUserContext(siteUrl, userName, password);

            // get web
            Web web = context.Web;
            context.Load(web, w => w.Url);
            context.ExecuteQuery();

            //check if the current web is a publishing web
            PublishingWeb pubWeb = PublishingWeb.GetPublishingWeb(context, web);
            context.Load(pubWeb);
            context.ExecuteQuery();

            if (pubWeb != null)
            {
                //retrieve the pages list
                List pagesList = pubWeb.Web.Lists.GetByTitle("Pages");
                context.Load(pubWeb);
                context.ExecuteQuery();

                // build CAML query to get all items in the Pages library
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope='RecursiveAll'></View>";

                //get all documents in the list
                ListItemCollection collListItem = pagesList.GetItems(camlQuery);

                context.Load(collListItem,
                items => items.Include(
                   item => item.Id,
                   item => item["Title"],
                   item => item["FileRef"])); //FileRef is the relative Url of the page
                context.ExecuteQuery();

                // get navigation terms
                var navigationTermSet = TaxonomyNavigation.GetTermSetForWeb(context, web, "GlobalNavigationTaxonomyProvider", true);
                var allNavigationTerms = navigationTermSet.GetAllTerms();

                context.Load(allNavigationTerms, t => t.Include(
                    i => i.TargetUrl,
                    i => i.LinkType,
                    i => i.TaxonomyName,
                    i => i.Parent));

                context.ExecuteQuery();

                // loop thru' all items in the 'Pages' library
                if (collListItem.Count > 0)
                {
                    foreach (ListItem oListItem in collListItem)
                    {
                        //get list item
                        context.Load(oListItem);
                        context.ExecuteQuery();


                        var navigationTermsForPage = allNavigationTerms.Where(
                            t => t.LinkType == NavigationLinkType.FriendlyUrl &&
                            t.TargetUrl.Value.Contains(oListItem["FileRef"].ToString()));

                        //context.Web.EnsureProperty(w => w.Url);

                        foreach (var navTerm in navigationTermsForPage)
                        {
                            var pageUrl = "";

                            context.Load(navTerm);

                            pageUrl = InsertUrlRecursive(navTerm, pageUrl);

                            Console.WriteLine($"{web.Url}{pageUrl}");
                        }
                    }
                }

            }
        }

        private static string InsertUrlRecursive(NavigationTerm navTerm, string pageUrl)
        {
            pageUrl = pageUrl.Insert(0, $"/{navTerm.TaxonomyName}");

            if (navTerm.Parent.ServerObjectIsNull == false)
            {
                pageUrl = InsertUrlRecursive(navTerm.Parent, pageUrl);
            }
            return pageUrl;
        }



        static ClientContext GetUserContext(string siteUrl, string userName, string password)
        {
            try
            {

                var securePassword = new SecureString();
                foreach (char c in password)
                {
                    securePassword.AppendChar(c);
                }
                var spoCredentials = new SharePointOnlineCredentials(userName, securePassword);
                var spoContext = new ClientContext(siteUrl);
                // Add User Agent information to context object to avoid throttling, 
                //  per guidance here -> https://docs.microsoft.com/en-us/sharepoint/dev/general-development/how-to-avoid-getting-throttled-or-blocked-in-sharepoint-online#how-to-decorate-your-http-traffic-to-avoid-throttling
                spoContext.ExecutingWebRequest += delegate (object sender, WebRequestEventArgs e)
                {
                    e.WebRequestExecutor.WebRequest.UserAgent = "NONISV|KiranKakanur|SPOUtilities/1.0";
                };
                spoContext.Credentials = spoCredentials;
                return spoContext;
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        static void EnableVersioningInDocLib(string siteUrl, string userName, string userPassword)
        {
            try
            {
                // get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                // get web
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                // get list
                List documentslist = web.Lists.GetByTitle("Documents");
                context.Load(documentslist);
                context.ExecuteQuery();

                // enable versioning if it isn't enabled
                if (!documentslist.EnableVersioning) { documentslist.EnableVersioning = true; documentslist.Update(); context.ExecuteQuery(); }

                // enable minor versions if it isn't enabled
                if (!documentslist.EnableMinorVersions) { documentslist.EnableMinorVersions = true; ; documentslist.Update(); context.ExecuteQuery(); }

                // set 5 major versions and retain drafts for 2 major versions
                if (documentslist.EnableVersioning && documentslist.EnableMinorVersions)
                {
                    documentslist.MajorVersionLimit = 5;
                    documentslist.MajorWithMinorVersionsLimit = 2;
                    documentslist.Update();
                    context.ExecuteQuery();
                }
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        static void DeleteOldDocumentVersions(string siteUrl, string userName, string userPassword, string csvFilePath)
        {
            try
            {
                int maxDocumentVersions = 15;//this is the maximum number of document versions that will be stored in the document library

                //get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                //get web
                Web web = context.Site.RootWeb;
                context.Load(web, w => w.Webs); //include subwebs
                context.ExecuteQuery();

                //caml query to get all documents in the library
                CamlQuery camlQuery = new CamlQuery();
                camlQuery.ViewXml = "<View Scope='RecursiveAll'></View>";

                //create StringBuilder object for storing documents info
                StringBuilder docsCSVFile = new StringBuilder();
                docsCSVFile.AppendLine("Title, URL, Versions");

                // Loop through all the webs  
                foreach (Web subWeb in web.Webs)
                {
                    context.Load(subWeb, subw => subw.Webs); //include title and all webs
                    context.ExecuteQuery();
                    Console.WriteLine("Processing site -> {0}", subWeb.Url);

                    // Loop through all webs in subwebs
                    foreach (Web subsubWeb in subWeb.Webs)
                    {
                        context.Load(subsubWeb, subsubw => subsubw.Title); //include title
                        context.ExecuteQuery();
                        Console.WriteLine("Processing site -> {0}", subsubWeb.Url);

                        //get list
                        List list = subsubWeb.Lists.GetByTitle("Documents");
                        context.Load(list);
                        context.ExecuteQuery();

                        //get all documents in the list
                        ListItemCollection collListItem = list.GetItems(camlQuery);

                        context.Load(collListItem,
                        items => items.Include(
                           item => item.Id,
                           item => item["Title"],
                           item => item["FileRef"]));
                        context.ExecuteQuery();

                        if (collListItem.Count > 0)
                        {
                            foreach (ListItem oListItem in collListItem)
                            {
                                //get list item
                                context.Load(oListItem, listItem => listItem.File, listItem => listItem.FileSystemObjectType);
                                context.ExecuteQuery();

                                //get file associated with list item
                                Microsoft.SharePoint.Client.File file = oListItem.File;
                                context.Load(file, f => f.CheckOutType, f => f.Versions);
                                context.ExecuteQuery();

                                if (oListItem.FileSystemObjectType == FileSystemObjectType.File) //check if it is a File. 
                                {
                                    if (file.CheckOutType == CheckOutType.None) //only process the file if it is not checked out
                                    {
                                        //get file version collection
                                        FileVersionCollection fvCollection = file.Versions;
                                        context.Load(fvCollection);
                                        context.ExecuteQuery();
                                        Console.WriteLine("Versions count for file {0} is {1} ", file.Name, file.Versions.Count.ToString());

                                        //check if current number of versions is greater than the max number of versions
                                        if (file.Versions.Count + 1 > maxDocumentVersions) //add 1 to versions count, since it is zero index based
                                        {
                                            Console.WriteLine("Processing document titled, {0}, with URL {1} and {2} versions", oListItem["Title"], oListItem["FileRef"], (file.Versions.Count + 1).ToString());
                                            docsCSVFile.AppendLine(oListItem["Title"] + "," + oListItem["FileRef"] + "," + (file.Versions.Count + 1).ToString());
                                            int versionsToDelete = (file.Versions.Count + 1) - maxDocumentVersions; //add 1 to versions count, since it is zero index based

                                            //loop through collection and delete old versions
                                            for (int i = 0; i < versionsToDelete; i++)
                                            {
                                                FileVersion fileVersion = fvCollection[0];//always delete the first (i.e. oldest) verion
                                                Console.WriteLine("fileVersion with label {0} created on {1} {2} will be deleted", fileVersion.VersionLabel, fileVersion.Created.ToShortDateString(), fileVersion.Created.ToShortTimeString());
                                                fileVersion.DeleteObject();
                                            }
                                            oListItem.SystemUpdate();
                                            context.ExecuteQuery();
                                        }//check for number of document versions
                                    }//check for file checkout
                                }//check if it is a file and not a folder
                            }//loop through documents 
                        }//check if library has documents
                    }//loop through subWeb.Webs
                }//loop through web.webs

                //Write docs info to CSV file
                System.IO.File.AppendAllText(csvFilePath, docsCSVFile.ToString());
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        static void GetLastModifiedInfo(string siteUrl, string userName, string userPassword, string csvFilePath)
        {
            try
            {
                //get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                //get web
                Web web = context.Site.RootWeb;
                context.Load(web, w => w.Webs); //include subwebs
                context.ExecuteQuery();

                //create StringBuilder object for storing documents info
                StringBuilder lastModCSVFile = new StringBuilder();
                lastModCSVFile.AppendLine("Title, URL, Site-LastItemModifiedDate, Site-LastItemUserModifiedDate, Lib-LastItemDeletedDate, Lib-LastItemModifiedDate, Lib-LastItemUserModifiedDate");

                // Loop through all the webs  
                foreach (Web subWeb in web.Webs)
                {
                    context.Load(subWeb, subw => subw.Webs); //include title and all webs
                    context.ExecuteQuery();

                    // Loop through all webs in subwebs
                    foreach (Web subsubWeb in subWeb.Webs)
                    {
                        context.Load(subsubWeb, subsubw => subsubw.Title, subsubw => subsubw.LastItemModifiedDate, subsubw => subsubw.LastItemUserModifiedDate); //include title, lastitemmodifieddate, lastitemusermodifieddate
                        context.ExecuteQuery();

                        Console.WriteLine("Processing site -> {0}", subsubWeb.Url);

                        //get list
                        List list = subsubWeb.Lists.GetByTitle("Documents");
                        context.Load(list, l => l.LastItemDeletedDate, l => l.LastItemModifiedDate, l => l.LastItemUserModifiedDate); //include lastitemdeleteddate, lastitemmodifieddate, lastitemusermodifieddate
                        context.ExecuteQuery();

                        //store info in CSV file
                        lastModCSVFile.AppendLine(subsubWeb.Title + "," +
                                                             subsubWeb.Url + "," +
                                                             subsubWeb.LastItemModifiedDate.ToShortDateString() + " " +
                                                             subsubWeb.LastItemModifiedDate.ToShortTimeString() + "," +
                                                             subsubWeb.LastItemUserModifiedDate.ToShortDateString() + " " +
                                                             subsubWeb.LastItemUserModifiedDate.ToShortTimeString() + "," +
                                                             list.LastItemDeletedDate.ToShortDateString() + " " +
                                                             list.LastItemDeletedDate.ToShortTimeString() + "," +
                                                             list.LastItemModifiedDate.ToShortDateString() + " " +
                                                             list.LastItemModifiedDate.ToShortTimeString() + "," +
                                                             list.LastItemUserModifiedDate.ToShortDateString() + " " +
                                                             list.LastItemUserModifiedDate.ToShortTimeString());
                    }//loop through subWeb.Webs
                }//loop through web.webs

                //Write docs info to CSV file
                System.IO.File.AppendAllText(csvFilePath, lastModCSVFile.ToString());
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        static void AddListPropertyBag(string siteUrl, string userName, string userPassword)
        {
            try
            {
                //get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                //get web
                Web web = context.Web;
                context.Load(web);
                context.ExecuteQuery();

                //get list
                List list = web.Lists.GetByTitle("Documents"); //the name of the list where you want to set the property bag values
                context.Load(list, olist => olist.RootFolder, olist => olist.RootFolder.Properties);
                context.ExecuteQuery();

                // get property bag
                var listProperties = list.RootFolder.Properties;

                //check if property bag that we are setting exists. if no, set it. if yes, delete existing and add it
                var pValue1 = list.RootFolder.Properties["IsSharePointOnlineAwesome"]; // get property value
                if (pValue1 == null) { listProperties["IsSharePointOnlineAwesome"] = "Yes"; } // set property value

                var pValue2 = list.RootFolder.Properties["DoYouLiveCloseToTheMoon"]; //get property value
                if (pValue2 == null) { listProperties["DoYouLiveCloseToTheMoon"] = "No"; } // set property value

                list.RootFolder.Update();
                context.ExecuteQuery();

                //read the property bag value
                pValue1 = list.RootFolder.Properties["IsSharePointOnlineAwesome"]; // get property value
                pValue2 = list.RootFolder.Properties["DoYouLiveCloseToTheMoon"]; // get property value
                Console.WriteLine("IsSharePointOnlineAwesome = {0}", pValue1);
                Console.WriteLine("DoYouLiveCloseToTheMoon = {0}", pValue2);
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }

        static void GetUserProfileProperties(string siteUrl, string userName, string userPassword)
        {
            try
            {
                // get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                // user account name in Claims format
                string accountName = "i:0#.f|membership|kiran@durianland.onmicrosoft.com";

                // skills values
                List<string> skills = new List<string>();
                skills.Add("SharePoint");
                skills.Add("CSOM");
                skills.Add("JavaScript");


                // Get the PeopleManager object and then get the target user's properties.
                PeopleManager peopleManager = new PeopleManager(context);

                // set single value profile property
                peopleManager.SetSingleValueProfileProperty(accountName, "AboutMe", "I love SharePoint!");

                // set multi value profile property
                peopleManager.SetMultiValuedProfileProperty(accountName, "SPS-Skills", skills);

                context.ExecuteQuery();

                // Get properties
                PersonProperties personProperties = peopleManager.GetPropertiesFor(accountName);

                // properties of the personProperties object.
                context.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
                context.ExecuteQuery();

                foreach (var property in personProperties.UserProfileProperties)
                {
                    Console.WriteLine(string.Format("{0}: {1}",
                        property.Key.ToString(), property.Value.ToString()));
                }

            }

            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }



        }

        // NOTE: The user account referenced in the SPOUserName setting in App.Config file must be a Term Store Administrator 
        // in order to to create Terms in Term Store
        // This is granted in the SharePoint Admin site (for example:https://<yourtenant>-admin.sharepoint.com/_layouts/15/termstoremanager.aspx)
        static void CreateTermsInTermStore(string siteUrl, string userName, string userPassword)
        {
            try
            {
                //get client context
                ClientContext context = GetUserContext(siteUrl, userName, userPassword);

                TaxonomySession taxSession = TaxonomySession.GetTaxonomySession(context);
                context.Load(taxSession);
                context.ExecuteQuery();


                if (taxSession != null)
                {
                    TermStore termStore = taxSession.GetDefaultSiteCollectionTermStore();
                    if (termStore != null)
                    {
                        //
                        // Create group, termset, and terms.
                        //
                        TermGroup myGroup = termStore.CreateGroup("Custom", Guid.NewGuid());
                        TermSet myTermSet = myGroup.CreateTermSet("Countries", Guid.NewGuid(), 1033);
                        myTermSet.CreateTerm("United States of America", 1033, Guid.NewGuid());
                        myTermSet.CreateTerm("Canada", 1033, Guid.NewGuid());
                        myTermSet.CreateTerm("Mexico", 1033, Guid.NewGuid());
                        myTermSet.CreateTerm("India", 1033, Guid.NewGuid());
                        myTermSet.CreateTerm("Thailand", 1033, Guid.NewGuid());
                        myTermSet.CreateTerm("Australia", 1033, Guid.NewGuid());

                        context.ExecuteQuery();
                    }
                }
            }
            catch (ClientRequestException clientEx)
            {
                Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
                throw clientEx;
            }
            catch (ServerException serverEx)
            {
                Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
                throw serverEx;
            }
        }


        //static void WorkWithManagedProperties(string siteUrl, string userName, string userPassword)
        //{
        //    try
        //    {
        //        //get client context
        //        ClientContext context = GetUserContext(siteUrl, userName, userPassword);

        //        //get Site
        //        Site site = context.Site;

        //        //Microsoft.SharePoint.Client.Search.Administration.SearchObjectLevel.Ssa


        //        SearchServiceApplication ssa = new SearchServiceApplication();
        //        SearchServiceApplicationProxy ssaproxy = new SearchServiceApplicationProxy();


        //        Guid guidTenant = new Guid("8d66b737-01b6-4190-907e-b5da9591faae"); //follow steps in this blog http://www.ktskumar.com/2017/01/access-sharepoint-online-using-postman/ to get tenant Id from your SPO environment
        //        Guid guidSiteCollection = new Guid("34b15b51-1499-4bf2-83e1-bfbbcd84a9ed"); //run this to get the site collection GUID https://<yourSPOSite>.sharepoint.com/sites/<sitename>/_api/Site from your SPO environment
        //        //Schema schema = new Schema(ssa, guidTenant, guidSiteCollection);


        //        SearchContext contextSearch = SearchContext.GetContext(ssa);
        //        Schema schema = new Schema(contextSearch);

        //        //Schema sspSchema = new Schema(SearchContext.GetContext(ssa, false));

        //        ManagedPropertyCollection collManagedProperties = schema.AllManagedProperties;

        //        ManagedProperty mpViewsLastMonths1 = collManagedProperties["viewslastmonths1"];

        //    }
        //    catch (ClientRequestException clientEx)
        //    {
        //        Console.WriteLine("Client side error occurred: {0} \n{1} " + clientEx.Message + clientEx.InnerException);
        //        throw clientEx;
        //    }
        //    catch (ServerException serverEx)
        //    {
        //        Console.WriteLine("Server side error occurred: {0} \n{1} " + serverEx.Message + serverEx.InnerException);
        //        throw serverEx;
        //    }
        //}
    }

}


