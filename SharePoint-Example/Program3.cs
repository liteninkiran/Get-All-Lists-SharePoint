using System;
using System.Security;
using System.IO;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace SharePoint_Example
{
    class Program3
    {
        static void Main()
        {
            // Define the site URL
            string siteUrl = "https://sharepoint.com";

            // Define credentials
            string userName = "";
            string passWord = "";

            // Set the Client Context
            ClientContext clientContext = GetClient(siteUrl, userName, passWord);

            // Get the SharePoint web  
            Web web = clientContext.Web;

            // Execute the query to the server  
            clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url);
            clientContext.ExecuteQuery();

            Guid listGuid = new Guid("ed2afef5-6f3b-497f-802e-bcadccaef8a2");

            SP.List oList = web.Lists.GetById(listGuid);
            SP.ListItem oListItem = oList.GetItemById(3317);

            // Load list & fields
            clientContext.Load(oList);
            clientContext.Load(oListItem);
            clientContext.ExecuteQuery();

            oListItem["Title"] = "J6781";

            Console.WriteLine(oListItem.FieldValues["Title"]);

            oListItem.Update();
            //oList.Update();
            clientContext.ExecuteQuery();

            Console.WriteLine(oListItem.FieldValues["Title"]);

            // Await user response
            Console.ReadLine();
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
    }
}
