using System;
using System.IO;
using System.Configuration;
using Microsoft.SharePoint.Client;
using SP = Microsoft.SharePoint.Client;
using System.Security;

namespace MoveListItemAttachments
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //Get site URL and credentials values from config
                Uri siteUri = new Uri(ConfigurationManager.AppSettings["SourceSite"].ToString());

                //Connect to SharePoint Online 
                using (ClientContext clientContext = new ClientContext(siteUri.ToString()))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in ConfigurationManager.AppSettings["DestinationPassword"].ToCharArray()) passWord.AppendChar(c);
                    clientContext.Credentials = new SharePointOnlineCredentials("vgouldla@ingramcontent.com", passWord);

                    if (clientContext != null)
                    {
                        //Source list
                        List sourceList = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["SourceList"]);
                        //Destination library
                        List destinationLibrary = clientContext.Web.Lists.GetByTitle(ConfigurationManager.AppSettings["DestinationLibrary"]);

                        // try to get all the list items
                        // could get in sections if it exceeds List View Threshold
                        CamlQuery camlQuery = new CamlQuery();
                        camlQuery.ViewXml = "<View><Query><OrderBy><FieldRef Name='Title' /></OrderBy></Query></View>";

                        ListItemCollection listItems = sourceList.GetItems(camlQuery);
                        FieldCollection listFields = sourceList.Fields;
                        clientContext.Load(sourceList);
                        clientContext.Load(listFields);
                        clientContext.Load(listItems);
                        clientContext.ExecuteQuery();

                        // Download attachments for each list item and then upload to new list item
                        foreach (ListItem item in listItems)
                        {
                            string attachmentURL = siteUri + "/Lists/" + ConfigurationManager.AppSettings["SourceList"].ToString() + "/Attachments/" + item["ID"];
                            Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(attachmentURL);
                            clientContext.Load(folder);

                            try
                            {
                                clientContext.ExecuteQuery();
                            }
                            catch (ServerException ex)
                            {
                                Console.WriteLine(ex.Message);
                                Console.WriteLine("No Attachment for ID " + item["ID"].ToString());
                            }

                            FileCollection attachments = folder.Files;
                            clientContext.Load(attachments);
                            clientContext.ExecuteQuery();

                            // write each file to local disk
                            foreach (SP.File file in folder.Files)
                            {
                                if (clientContext.HasPendingRequest)
                                {
                                    clientContext.ExecuteQuery();
                                }
                                var fileRef = file.ServerRelativeUrl;
                                var fileInfo = SP.File.OpenBinaryDirect(clientContext, fileRef);

                                using (var memory = new MemoryStream())
                                {
                                    byte[] buffer = new byte[1024 * 64];
                                    int nread = 0;
                                    while ((nread = fileInfo.Stream.Read(buffer, 0, buffer.Length)) > 0)
                                    {
                                        memory.Write(buffer, 0, nread);
                                    }
                                    memory.Seek(0, SeekOrigin.Begin);
                                    // at this point you have the contents of your file in memory
                                    // save to computer
                                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(clientContext, string.Format("/{0}/{1}", ConfigurationManager.AppSettings["AttachmentLibrary"], System.IO.Path.GetFileName(file.Name)), memory, true);
                                }

                                // this call avoids potential problems if any requests are still pending
                                if (clientContext.HasPendingRequest)
                                {
                                    clientContext.ExecuteQuery();
                                }

                                SP.File newFile = clientContext.Web.GetFileByServerRelativeUrl(string.Format("/{0}/{1}", ConfigurationManager.AppSettings["AttachmentLibrary"], System.IO.Path.GetFileName(file.Name)));
                                clientContext.Load(newFile);
                                clientContext.ExecuteQuery();

                                //check out to make sure not to create multiple versions
                                newFile.CheckOut();

                                FieldLookupValue applicationName = item["Source"] as FieldLookupValue;
                                
                                // app name may be null
                                if (applicationName == null) applicationName = new FieldLookupValue();

                                applicationName.LookupId = Convert.ToInt32(item["ID"]);
                                ListItem newItem = newFile.ListItemAllFields;
                                newItem["From_x0020_Source"] = applicationName;
                                newItem.Update();

                                // use OverwriteCheckIn type to make sure not to create multiple versions 
                                newFile.CheckIn(string.Empty, CheckinType.OverwriteCheckIn);

                                // Clear requests if any if pending
                                if (clientContext.HasPendingRequest)
                                {
                                    clientContext.ExecuteQuery();
                                }
                            }
                            Console.WriteLine("All list items and attachments copied over. Press any key to close");
                            Console.ReadKey();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed: " + ex.Message);
                Console.WriteLine("Stack Trace: " + ex.StackTrace);
                Console.ReadKey();
            }
        }
    }
}