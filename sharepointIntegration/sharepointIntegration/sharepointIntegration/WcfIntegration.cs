/*
 * Use WCF and LINQ to Sharepoint to integrate sharepoint
 * Author: Alessandro Graps
 * Year: 2013
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using sharepointIntegration.ServiceReference1;

namespace sharepointIntegration
{
    public static class SharePointWcfService
    {
        /// <summary>
        /// Gets the item id.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public static ListaTestDataVerificationItem GetItemId(string name)
        {        
            try
            {
                DevelopementDataContext dataContext =
                new DevelopementDataContext(new Uri("http://sharepoint.site.com/_vti_bin/ListData.svc/"));

                dataContext.Credentials = CredentialCache.DefaultNetworkCredentials;

                var result = dataContext.ListaTestDataVerification.ToList();
                return result.Where(o => o.Title == name).FirstOrDefault();

            }
            catch (Exception ex)
            {
                throw;
            }
            finally
            {                
            }
        }

        /// <summary>
        /// Uploads the document by agent.
        /// </summary>
        /// <param name="nameItem">The name item.</param>
        public static void UploadDocumentByAgent(string nameItem)
        {            
            DevelopementDataContext dataContext =
                 new DevelopementDataContext(new Uri("http://sharepoint.site.com/_vti_bin/ListData.svc/"));

            dataContext.Credentials = CredentialCache.DefaultNetworkCredentials;

            ListaTestDataVerificationItem item = GetItemId(nameItem);

            if (item != null)
            {
                var attachment = new AttachmentsItem { EntitySet = "ListaTestDataVerification", ItemId = item.Id, Name = "filetoUpload.ext" };
                dataContext.AddToAttachments(attachment);

                FileStream stream = new FileStream(@"c:\temp\testo.txt", FileMode.Open, FileAccess.Read);

                dataContext.SetSaveStream(attachment, stream, false, "Item", attachment.EntitySet + "|" + attachment.ItemId + "|" + attachment.Name);
                dataContext.SaveChanges();      
            }                                  
        }

        /// <summary>
        /// Creates the new item.
        /// </summary>
        /// <param name="title">The title.</param>
        public static void CreateNewItem(string title)
        {
            DevelopementDataContext dataContext =
                 new DevelopementDataContext(new Uri("http://sharepoint.site.com/_vti_bin/ListData.svc/"));

            dataContext.Credentials = CredentialCache.DefaultNetworkCredentials;

            string contentType = "Item";
            ListaTestDataVerificationItem newItem = new ListaTestDataVerificationItem();
            newItem.ContentType = contentType;
            newItem.Title = "AA3A6A5351";
            newItem.IdAgent = "1fr2334";
            newItem.IdCustomer = "23344";
            newItem.IdBackOfficeOperator = "2344";
            newItem.AgentUsername = "32344";
            
            dataContext.AddToListaTestDataVerification(newItem);
        }

        /// <summary>
        /// Gets all item.
        /// </summary>
        /// <returns></returns>
        public static List<ListaTestDataVerificationItem> GetAllItem()
        {
            DevelopementDataContext dataContext =
                   new DevelopementDataContext(new Uri("http://sharepoint.site.com/_vti_bin/ListData.svc/"));

            List<ListaTestDataVerificationItem> list = new List<ListaTestDataVerificationItem>();

            dataContext.Credentials = CredentialCache.DefaultNetworkCredentials;
            var result = dataContext.ListaTestDataVerification.ToList();
            foreach(var item in result)
            {
                list.Add(item);
            }

            return list;
        }

        /// <summary>
        /// Moves the item from list to another list.
        /// </summary>
        public static void MoveItemFromListToAnotherList()
        {
            DevelopementDataContext dataContext =
                  new DevelopementDataContext(new Uri("http://sharepoint.site.com/_vti_bin/ListData.svc"));

            List<ListaTestDataVerificationItem> list = new List<ListaTestDataVerificationItem>();

            dataContext.Credentials = CredentialCache.DefaultNetworkCredentials;
            var result = dataContext.ListaTestDataVerification.ToList();
            foreach (var item in result)
            {
                list.Add(item);
            }

            ListaTestDataVerificationItem it = list.First(o => o.Id == 11);
            

            ObjectModelIntegration objectModel = new ObjectModelIntegration();
            objectModel.DownloadAttachedFileFromListItem(@"http://http://sharepoint.site.com",
                    @"developement", it.Id, it.AttachmentsName, "ListaTestDataVerification", @"C:\temp");  
            
        
        }
    }
}
