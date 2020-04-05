using System;
using System.Security;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace SharePoint_Example
{
    class Program
    {
        static void Main()
        {
            // Define the site URL
            string siteUrl = "https://mysite.sharepoint.com/";

            // Define credentials
            string userName = "person@domain.com";
            string passWord = "my-password";

            // Define object names
            string listName = "List Name";
            string fieldName = "Field Name";
            int itemId = 1;

            // Set the Client Context
            ClientContext clientContext = GetClient(siteUrl, userName, passWord);

            // Find last status change
            FindStatusChange(clientContext, listName);

            // Await user response
            Console.ReadLine();

            // Exit procedure
            return;

            // Loop through all versions of an item in a list
            LoopThroughAllVersions(clientContext, listName, itemId);

            // Loop through all fields in a list
            LoopThroughAllFields(clientContext, listName);

            // Loop through all versions of an item in a list
            LoopThroughAllVersions(clientContext, listName, itemId);

            // Loop through all lists
            LoopThroughAllLists(clientContext);

            // Loop through all fields in specified list
            FieldLoop(clientContext, listName);

            // Find all lists with specified field name
            FindLists(clientContext, fieldName);
        }

        static void FindStatusChange(ClientContext clientContext, string listName)
        {
            // Find the specified list
            SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

            // Create a new Caml Query
            CamlQuery camlQuery = new CamlQuery();

            // Set the XML
            camlQuery.ViewXml = "<View><Query>{0}</Query></View>";

            // Define a collection to store the list items in
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            // Load in the items
            clientContext.Load(collListItem);
            clientContext.ExecuteQuery();

            // Store field names
            string jobField = "Title";
            string statusField = "Status";
            string statusValue = "Commissioned";

            // Create a blank version object
            SP.ListItemVersion commVersion = null;

            // Loop through each item in the collection
            foreach (ListItem oListItem in collListItem)
            {
                // Load the Versions
                clientContext.Load(oListItem.Versions);
                clientContext.ExecuteQuery();

                // Store details of the item
                string job = oListItem[jobField].ToString().Trim();
                string status = oListItem[statusField].ToString().Trim();

                // Store output strings
                string created = null;
                string versionLabel = null;
                string versionCount = oListItem.Versions.Count.ToString();

                // Find the earliest version where Status has been set to our chosen value
                commVersion = CheckHistory(clientContext, oListItem, statusField, statusValue);

                // If we found a version, store the details
                if (commVersion != null)
                {
                    created = commVersion.Created.ToString();
                    versionLabel = commVersion.VersionLabel;
                }

                Console.WriteLine("{0},{1},{2},{3},{4}", job, created, versionLabel, versionCount, status);
            }
        }
        
        public static SP.ListItemVersion CheckHistory(ClientContext clientContext, SP.ListItem oListItem, string statusField, string statusValue)
        {
            // Create a blank object to return
            SP.ListItemVersion commVersion = null;

            // Store the status change date
            DateTime commDate = new DateTime();

            // Loop through each Version
            foreach (SP.ListItemVersion versionItem in oListItem.Versions)
            {
                // Store the status value
                string status = versionItem[statusField].ToString();

                // If we have a matching status, then store the earliest version
                if(status == statusValue)
                {
                    // Check against this version's create date. Dates start on year 0001, so check that too.
                    if(versionItem.Created < commDate || commDate.Year == 1)
                    {
                        // We found an earlier version
                        commDate = versionItem.Created;
                        commVersion = versionItem;
                    }
                }
            }

            // Return the earliest version where the status value matches the supplied value
            return commVersion;
        }

        static void LoopThroughAllVersions(ClientContext clientContext, string listName, int id)
        {
            // Find the specified list
            SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

            // Find the specified item
            SP.ListItem oListItem = oList.GetItemById(id);

            // Load the Item
            clientContext.Load(oListItem);

            // Load the Versions
            clientContext.Load(oListItem.Versions);

            // Do this bit
            clientContext.ExecuteQuery();

            // Print out the item identifier
            Console.WriteLine(oListItem["Title"] + " Previous Versions");
            Console.WriteLine("");

            // Loop through each Version
            foreach(SP.ListItemVersion versionItem in oListItem.Versions)
            {
                // Output something about the version
                Console.WriteLine("VersionLabel: " + versionItem.VersionLabel);
                Console.WriteLine("IsCurrentVersion: " + versionItem.IsCurrentVersion);
                Console.WriteLine("Created: " + versionItem.Created);
                Console.WriteLine("CreatedBy: " + versionItem.CreatedBy);
                Console.WriteLine("");

                // Retrieve the fields
                SP.FieldCollection collField = versionItem.Fields;

                clientContext.Load(collField);
                clientContext.ExecuteQuery();

                int i = 0;

                // Loop through fields
                foreach(SP.Field oField in collField)
                {
                    string fieldName = oField.Title;
                    string fieldValue = "null";

                    try
                    {
                        if (oListItem[oField.InternalName] != null)
                        {
                            fieldValue = versionItem[oField.InternalName].ToString();
                        }
                    }
                    catch(Exception e)
                    {
                        fieldName = oField.InternalName;
                        fieldValue = "Error";
                    }
                    finally
                    {
                        Console.WriteLine("{0}) {1}: {2}", i, fieldName, fieldValue);
                        i++;
                    }
                }

                Console.WriteLine("");
                Console.ReadLine();
            }
        }

        static void FindLists(ClientContext clientContext, string fieldName)
        {
            // Define website object
            Web oWebsite = clientContext.Web;

            // Store lists in a collection
            ListCollection collList = oWebsite.Lists;

            // Retreive lists
            clientContext.Load(collList);
            clientContext.ExecuteQuery();

            // Loop through each list in the collection
            foreach (SP.List oList in collList)
            {
                // Store the internal field name
                string fieldNameIntl = HasField(clientContext, oList, fieldName);

                // Output list name
                if(fieldNameIntl != "")
                {
                    Console.WriteLine("List Name: {0} {1} ({2})", oList.Title, fieldName, fieldNameIntl);
                }
            }
        }

        static void LoopThroughAllFields(ClientContext clientContext, string listName)
        {
            // Find the specified list
            SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

            // Retrieve the fields
            SP.FieldCollection collField = oList.Fields;

            // Load list & fields
            clientContext.Load(collField);
            clientContext.ExecuteQuery();

            int i = 0;

            // Loop through fields
            foreach (SP.Field oField in collField)
            {
                i++;

                Console.WriteLine("{0} Title         : {1}", i, oField.Title);
                Console.WriteLine("{0} Internal Name : {1}", i, oField.InternalName);
                Console.WriteLine("{0} Type          : {1}", i, oField.TypeAsString);
                Console.WriteLine("");
            }
        }

        public static string HasField(ClientContext clientContext, SP.List oList, string fieldName)
        {
            // Retrieve the fields
            SP.FieldCollection collField = oList.Fields;

            // Load list & fields
            clientContext.Load(collField);
            clientContext.ExecuteQuery();

            // Loop through fields
            foreach (SP.Field oField in collField)
            {
                // Check if field's title matches the specified field name
                if(oField.Title == fieldName)
                {
                    // Return the internal name for use later on
                    return oField.InternalName;
                }
            }

            // Return an empty string if the field wasn't found
            return "";
        }

        static void LoopThroughAllLists(ClientContext clientContext)
        {
            // Define website object
            Web oWebsite = clientContext.Web;

            // Store lists in a collection
            ListCollection collList = oWebsite.Lists;

            // Retreive lists
            clientContext.Load(collList);
            clientContext.ExecuteQuery();

            // Initialise row counter
            int i = 0;

            // Loop through each list in the collection
            foreach (SP.List oList in collList)
            {
                // Increment counter
                i++;

                // Output list name
                Console.WriteLine("Row: {0} Title: {1} Created: {2}", i, oList.Title, oList.Created.ToString());

                // Loop through items in list
                LoopThroughAllItems(clientContext, oList.Title);

                // Only print first 10 lists
                if (i >= 10)
                {
                    break;
                }
            }
        }

        public static ClientContext GetClient(string siteUrl, string userName, string passWord)
        {
            // Create a Secure String for password
            SecureString secPassWord = new SecureString();

            // Append characters of password to Secure String
            foreach (char c in passWord.ToCharArray())
            {
                secPassWord.AppendChar(c);
            }

            // Set the Client Context
            ClientContext clientContext = new ClientContext(siteUrl);

            // Enter credentials
            clientContext.Credentials = new SharePointOnlineCredentials(userName, secPassWord);

            // Return the object
            return clientContext;
        }

        public static void LoopThroughAllItems(ClientContext clientContext, string listName)
        {
            // Retrieve the list
            SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

            // Create a new Caml Query
            CamlQuery camlQuery = new CamlQuery();

            // Set the XML
            camlQuery.ViewXml = "<View><Query>{0}</Query></View>";

            // Define a collection to store the list items in
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            // Load in the items
            clientContext.Load(collListItem);
            clientContext.ExecuteQuery();

            // Initialise row counter
            int i = 0;

            // Loop through each item in the collection
            foreach(ListItem oListItem in collListItem)
            {
                // Increment counter
                i++;

                // Output something from the list
                Console.WriteLine("    Row: {0} ID: {1} Title: {2}", i, oListItem.Id, oListItem["Title"]);

                // Only print first 10 records
                if (i >= 10)
                {
                    break;
                }
            }

            // Print out a line break
            Console.WriteLine("");
        }

        public static void FieldLoop(ClientContext clientContext, string listName)
        {
            // Retrieve the list
            SP.List oList = clientContext.Web.Lists.GetByTitle(listName);

            // Retrieve the fields
            SP.FieldCollection collField = oList.Fields;

            // Load list & fields
            clientContext.Load(oList);
            clientContext.Load(collField);
            clientContext.ExecuteQuery();

            // Loop through fields
            foreach(SP.Field oField in collField)
            {
                Console.WriteLine("List: {0} \n\t Field Title: {1} \n\t Field Internal Name: {2}", oList.Title, oField.Title, oField.InternalName);
            }

            // Print out a line break
            Console.WriteLine("");
        }

    }
}
