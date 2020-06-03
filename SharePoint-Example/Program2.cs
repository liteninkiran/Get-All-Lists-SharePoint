using System;
using System.Security;
using System.IO;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;

namespace SharePoint_Example
{
    class Program2
    {
        static void Main2()
        {
            // Define the site URL
            string siteUrl = "https://sharepoint.com";

            // Define credentials
            string userName = "";
            string passWord = "";

            // Set the Client Context
            ClientContext clientContext = GetClient(siteUrl, userName, passWord);

            string path = @"C:\Temp\xxx.txt";

            using (StreamWriter file = new StreamWriter(path))
            {
                GetJobProposals(clientContext, file);
            }

            Console.Write("Output to {0}", path);

            // Await user response
            Console.ReadLine();
        }

        private static void ListLoop(ClientContext clientContext, Web web, int i, string parent, StreamWriter file)
        {
            // Store lists in a collection
            ListCollection collList = web.Lists;

            // Retreive lists
            clientContext.Load(collList);
            clientContext.ExecuteQuery();

            // Initialise row counter
            int j = 0;

            // Loop through each list in the collection
            foreach (SP.List oList in collList)
            {
                // Increment counter
                j++;

                //OutputList(web, oList, i, j, parent);

                FieldLoop(clientContext, web, oList, i, j, parent, file);

                Console.WriteLine(oList.Title);

                //if (j >= 5){break;}
            }
        }

        public static void FieldLoop(ClientContext clientContext, Web web, SP.List oList, int i, int j, string parent, StreamWriter file)
        {
            // Retrieve the fields
            SP.FieldCollection collField = oList.Fields;

            // Load list & fields
            clientContext.Load(oList);
            clientContext.Load(collField);
            clientContext.ExecuteQuery();

            int k = 0;

            // Loop through fields
            foreach (SP.Field oField in collField)
            {
                k++;

                OutputField(web, oList, oField, i, j, k, parent, file);
            }
        }

        private static void OutputList(Web web, SP.List oList, int i, int j, string parent)
        {
            string output = "{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}";

            // Output list name
            Console.WriteLine
            (
                output,
                i,
                web.Title,
                parent,
                web.Url,
                j,
                oList.Id,
                oList.Title,
                oList.Description.Replace('\r', ' ').Replace('\n', ' '),
                oList.ItemCount,
                oList.EnableMinorVersions,
                oList.EnableVersioning,
                oList.MajorWithMinorVersionsLimit,
                oList.MajorVersionLimit
            );
        }

        private static void OutputField(Web web, SP.List oList, SP.Field oField, int i, int j, int k, string parent, StreamWriter file)
        {
            // Store the Field Type
            string fieldType = oField.TypeAsString;

            // We will exclude some entries by Field Type
            bool skip = false;

            // These 2 attributes will depend on the field type
            string maxLength = null;
            string formula = null;

            // Find field types for specific actions
            switch (fieldType)
            {
                // Max Length
                case "Text":

                    // Convert to "Text Field"
                    SP.FieldText textField = (SP.FieldText)oField;

                    // We now have access to the attribute
                    maxLength = textField.MaxLength.ToString();

                    // Break out of the switch
                    break;

                // Formula
                case "Calculated":

                    // Convert to "Calculated Field"
                    FieldCalculated calcField = (FieldCalculated)oField;

                    // We now have access to the attribute
                    formula = calcField.Formula;

                    // Break out of the switch
                    break;

                // We won't include "Computed" or "Lookup" fields
                case "Lookup":
                case "Computed":

                    // Skip
                    skip = true;

                    // Break out of the switch
                    break;
            }

            if (skip == false)
            {
                string fieldName = oField.Title.Replace('\r', ' ').Replace('\n', ' ');
                string defaultValue = oField.DefaultValue;
                string enforceUniqueValues = oField.EnforceUniqueValues.ToString();
                string required = oField.Required.ToString();
                string readOnly = oField.ReadOnlyField.ToString();
                string isSealed = oField.Sealed.ToString();

                string lineText =
                    i.ToString() + "|" +
                    web.Title + "|" +
                    parent + "|" +
                    web.Url + "|" +
                    j.ToString() + "|" +
                    oList.Id + "|" +
                    oList.Title + "|" +
                    k.ToString() + "|" +
                    fieldName + "|" +
                    fieldType + "|" +
                    maxLength + "|" +
                    oField.FromBaseType + "|" +
                    enforceUniqueValues + "|" +
                    required + "|" +
                    readOnly + "|" +
                    defaultValue + "|" +
                    formula;


                file.WriteLine(lineText);
            }
        }

        private static void GetJobProposals(ClientContext clientContext, StreamWriter file)
        {
            // Get the SharePoint web  
            Web web = clientContext.Web;

            // Execute the query to the server  
            clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url);
            clientContext.ExecuteQuery();

            Guid guid = new Guid("ed2afef5-6f3b-497f-802e-bcadccaef8a2");

            SP.List oList = web.Lists.GetById(guid);

            // Retrieve the fields
            SP.FieldCollection collField = oList.Fields;

            // Load list & fields
            clientContext.Load(oList);
            clientContext.Load(collField);
            clientContext.ExecuteQuery();

            // Create a new Caml Query
            CamlQuery camlQuery = new CamlQuery();

            // Set the XML
            camlQuery.ViewXml = "<View><Query>{0}</Query></View>";

            // Define a collection to store the list items in
            ListItemCollection collListItem = oList.GetItems(camlQuery);

            // Load in the items
            clientContext.Load(collListItem);
            clientContext.ExecuteQuery();

            // Loop through each item in the list
            foreach (ListItem oListItem in collListItem)
            {
                string status = oListItem.FieldValues["Status"].ToString();

                if (status == "Won")
                {
                    // Load the Versions
                    clientContext.Load(oListItem.Versions);
                    clientContext.ExecuteQuery();

                    string id = oListItem.FieldValues["ID"].ToString();
                    string title = oListItem.FieldValues["Title"].ToString();

                    foreach (SP.ListItemVersion versionItem in oListItem.Versions)
                    {
                        string created = versionItem.Created.ToString();
                        string versionStatus = versionItem.FieldValues["Status"].ToString();
                        string versionNumber = versionItem.VersionLabel;

                        string lineText = id + "|" + title + "|" + versionNumber + "|" + created + "|" + versionStatus;

                        file.WriteLine(lineText);
                    }
                }
            }
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
            //Console.WriteLine(oListItem["Title"] + " Previous Versions");
            //Console.WriteLine("");

            // Loop through each Version
            foreach (SP.ListItemVersion versionItem in oListItem.Versions)
            {
                // Output something about the version
                //Console.WriteLine("VersionLabel: " + versionItem.VersionLabel);
                //Console.WriteLine("IsCurrentVersion: " + versionItem.IsCurrentVersion);
                //Console.WriteLine("Created: " + versionItem.Created);
                //Console.WriteLine("CreatedBy: " + versionItem.CreatedBy);
                //Console.WriteLine("");

                // Retrieve the fields
                SP.FieldCollection collField = versionItem.Fields;

                clientContext.Load(collField);
                clientContext.ExecuteQuery();

                int i = 0;

                //Console.WriteLine("0,Version,{0}", versionItem.VersionLabel);

                // Loop through fields
                foreach (SP.Field oField in collField)
                {
                    string fieldName = oField.Title;
                    string fieldValue = "null";

                    if (oField.TypeAsString != "Computed")
                    {
                        try
                        {
                            if (oListItem[oField.InternalName] != null)
                            {
                                fieldValue = versionItem[oField.InternalName].ToString();
                            }
                        }
                        catch (Exception e)
                        {
                            //fieldName = oField.InternalName;
                            fieldValue = "Error";
                        }
                        finally
                        {
                            i++;
                            Console.WriteLine("{0},{1},{2},{3}", versionItem.VersionLabel, i, fieldName, '"' + fieldValue + '"');
                        }
                    }
                }

                //Console.WriteLine("");
                //Console.ReadLine();
            }
        }





        private static void GetSiteAndSubSites(ClientContext clientContext, bool recursive, StreamWriter file)
        {
            // Get the SharePoint web  
            Web web = clientContext.Web;

            // Execute the query to the server  
            clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url);
            clientContext.ExecuteQuery();

            int i = 1;

            ListLoop(clientContext, web, i, null, file);

            // Loop through sub-sites
            GetSubSites(clientContext, web, recursive, null, ref i, file);
        }

        private static void GetSubSites(ClientContext clientContext, Web web, bool recursive, string parent, ref int i, StreamWriter file)
        {
            //if (i > 5){return;}

            // Load objects
            clientContext.Load(web, website => website.Webs, website => website.Title, website => website.Url);

            // Execute the query to the server  
            clientContext.ExecuteQuery();

            // Loop through all the webs  
            foreach (Web subWeb in web.Webs)
            {
                // Check whether it is an app URL or not - If not then get into this block  
                if (subWeb.Url.Contains(web.Url))
                {
                    i++;

                    // Store full path of parent
                    string parentNew = null;

                    // If the incoming parent variable is null, we are at the top
                    if (parent != null)
                    {
                        // New parent is incoming parent with the (parent) web title
                        parentNew = parent + " --> " + web.Title;
                    }
                    else
                    {
                        // New parent is just the (parent) web title
                        parentNew = web.Title;
                    }

                    ListLoop(clientContext, subWeb, i, parentNew, file);

                    // Loop through sub-sites
                    if (recursive)
                    {
                        GetSubSites(clientContext, subWeb, recursive, parentNew, ref i, file);
                    }
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
    }
}
