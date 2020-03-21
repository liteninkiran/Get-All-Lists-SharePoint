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

            // Set the Client Context
            ClientContext clientContext = GetClient(siteUrl, userName, passWord);

            // Loop through all lists
            ListLoop(clientContext);

            // Await user response
            Console.ReadLine();
        }

        static void ListLoop(ClientContext clientContext)
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

            foreach (SP.List oList in collList)
            {
                // Increment counter
                i++;

                // Output list name
                Console.WriteLine("Row: {0} Title: {1} Created: {2}", i, oList.Title, oList.Created.ToString());

                // Loop through items in list
                ItemListLoop(clientContext, oList.Title);

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

            return clientContext;
        }

        public static void ItemListLoop(ClientContext clientContext, string listName)
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
                Console.WriteLine("    Row: {0} ID: {1} Title: {3}", i, oListItem.Id, oListItem["Title"]);

                // Only print first 10 records
                if(i >= 10)
                {
                    break;
                }
            }

            // Print out a line break
            Console.WriteLine("");
        }
    }
}
