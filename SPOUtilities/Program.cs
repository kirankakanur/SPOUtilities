using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;
using System.Configuration;
using System.IO;
using Microsoft.SharePoint.Client;

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
            if (string.IsNullOrEmpty(RunMode))
            {
                Console.WriteLine("1 - Enable Major and Minor Versions in Document Library");
                Console.WriteLine("2 - Delete Old Document Versions in Document Library");
                Console.WriteLine("3 - Get Last Modified Information from Site and Document Library");
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
                        EnableVersioningInDocLib(siteUrl, userName, password, csvFilePath);
                        break;

                    case "2":
                        DeleteOldDocumentVersions(siteUrl, userName, password, csvFilePath);
                        break;

                    case "3":
                        GetLastModifiedInfo(siteUrl, userName, password, csvFilePath);
                        break;
                }
            }
            else
            {
                Console.WriteLine("Please select an option and try again.");
                Console.ReadLine();
            }

        }


        static ClientContext GetUserContext(string siteUrl, string userName, string password)
        {
            var securePassword = new SecureString();
            foreach (char c in password)
            {
                securePassword.AppendChar(c);
            }

            var spoCredentials = new SharePointOnlineCredentials(userName, securePassword);
            var spoContext = new ClientContext(siteUrl);

            spoContext.Credentials = spoCredentials;
            return spoContext;

        }

        static void EnableVersioningInDocLib(string siteUrl, string userName, string userPassword, string csvFilePath)
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

                if (documentslist.EnableVersioning && !documentslist.EnableMinorVersions)
                {
                    documentslist.EnableMinorVersions = true;
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

        static void GetLastModifiedInfo(string siteUrl, string userName, string userPassword, string csvFilePath)
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
    }
}
