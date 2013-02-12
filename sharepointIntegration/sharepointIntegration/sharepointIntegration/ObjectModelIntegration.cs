/*
 * Use Object Model to integrate sharepoint
 * Author: Alessandro Graps
 * Year: 2013
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace sharepointIntegration
{
    public class ObjectModelIntegration
    {
        /// <summary>
        /// Creates the folder.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="relativePath">The relative path.</param>
        /// <param name="folderName">Name of the folder.</param>
        public void CreateFolder(string siteUrl, string listName, string relativePath, string folderName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                ListItemCreationInformation newItem = new ListItemCreationInformation();
                newItem.UnderlyingObjectType = FileSystemObjectType.Folder;
                newItem.FolderUrl = siteUrl + "/lists/" + listName;
                if (!relativePath.Equals(string.Empty))
                {
                    newItem.FolderUrl += "/" + relativePath;
                }
                newItem.LeafName = folderName;
                ListItem item = list.AddItem(newItem);
                item.Update();
                clientContext.ExecuteQuery();
            }
        }

        /// <summary>
        /// Deletes the folder.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="relativePath">The relative path.</param>
        /// <param name="folderName">Name of the folder.</param>
        public void DeleteFolder(string siteUrl, string listName, string relativePath, string folderName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                CamlQuery query = new CamlQuery();
                query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                "<Query>" +
                                    "<Where>" +
                                        "<And>" +
                                            "<Eq>" +
                                                "<FieldRef Name=\"FSObjType\" />" +
                                                "<Value Type=\"Integer\">1</Value>" +
                                             "</Eq>" +
                                              "<Eq>" +
                                                "<FieldRef Name=\"Title\"/>" +
                                                "<Value Type=\"Text\">" + folderName + "</Value>" +
                                              "</Eq>" +
                                        "</And>" +
                                     "</Where>" +
                                "</Query>" +
                                "</View>";

                if (relativePath.Equals(string.Empty))
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName;
                }
                else
                {
                    query.FolderServerRelativeUrl = "/lists/" + listName + "/" + relativePath;
                }

                var folders = list.GetItems(query);

                clientContext.Load(list);
                clientContext.Load(folders);
                clientContext.ExecuteQuery();
                if (folders.Count == 1)
                {
                    folders[0].DeleteObject();
                    clientContext.ExecuteQuery();
                }
            }
        }

        //public void RenameFolder(string siteUrl, string listName, string relativePath, string folderName, string folderNewName)
        //{
        //    using (ClientContext clientContext = new ClientContext(siteUrl))
        //    {
        //        Web web = clientContext.Web;
        //        List list = web.Lists.GetByTitle(listName);

        //        string FolderFullPath = GetFullPath(listName, relativePath, folderName);

        //        CamlQuery query = new CamlQuery();
        //        query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
        //                        "<Query>" +
        //                            "<Where>" +
        //                                "<And>" +
        //                                    "<Eq>" +
        //                                        "<FieldRef Name=\"FSObjType\" />" +
        //                                        "<Value Type=\"Integer\">1</Value>" +
        //                                     "</Eq>" +
        //                                      "<Eq>" +
        //                                        "<FieldRef Name=\"Title\"/>" +
        //                                        "<Value Type=\"Text\">" + folderName + "</Value>" +
        //                                      "</Eq>" +
        //                                "</And>" +
        //                             "</Where>" +
        //                        "</Query>" +
        //                        "</View>";

        //        if (relativePath.Equals(string.Empty))
        //        {
        //            query.FolderServerRelativeUrl = "/lists/" + listName;
        //        }
        //        else
        //        {
        //            query.FolderServerRelativeUrl = "/lists/" + listName + "/" + relativePath;
        //        }
        //        var folders = list.GetItems(query);

        //        clientContext.Load(list);
        //        clientContext.Load(list.Fields);
        //        clientContext.Load(folders, fs => fs.Include(fi => fi["Title"],
        //            fi => fi["DisplayName"],
        //            fi => fi["FileLeafRef"]));
        //        clientContext.ExecuteQuery();

        //        if (folders.Count == 1)
        //        {

        //            folders[0]["Title"] = folderNewName;
        //            folders[0]["FileLeafRef"] = folderNewName;
        //            folders[0].Update();
        //            clientContext.ExecuteQuery();
        //        }
        //    }
        //}

        /// <summary>
        /// Searches the folder.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="relativePath">The relative path.</param>
        public void SearchFolder(string siteUrl, string listName, string relativePath)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                Web web = clientContext.Web;
                List list = web.Lists.GetByTitle(listName);

                string FolderFullPath = null;

                CamlQuery query = CamlQuery.CreateAllFoldersQuery();

                if (relativePath.Equals(string.Empty))
                {
                    FolderFullPath = "/lists/" + listName;
                }
                else
                {
                    FolderFullPath = "/lists/" + listName + "/" + relativePath;
                }
                if (!string.IsNullOrEmpty(FolderFullPath))
                {
                    query.FolderServerRelativeUrl = FolderFullPath;
                }
                IList<Folder> folderResult = new List<Folder>();

                var listItems = list.GetItems(query);

                clientContext.Load(list);
                clientContext.Load(listItems, litems => litems.Include(
                    li => li["DisplayName"],
                    li => li["Id"]
                    ));

                clientContext.ExecuteQuery();

                foreach (var item in listItems)
                {

                    Console.WriteLine("{0}----------{1}", item.Id, item.DisplayName);
                }
            }
        }

        /// <summary>
        /// Uploads the file in library.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="libraryName">Name of the library.</param>
        /// <param name="subfolderPath">The subfolder path.</param>
        /// <param name="fileName">Name of the file.</param>
        public void UploadFileInLibrary(string siteUrl, string webName, string libraryName, string subfolderPath, string fileName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {

                string uploadLocation = Path.GetFileName(fileName);
                if (!string.IsNullOrEmpty(subfolderPath))
                {
                    uploadLocation = string.Format("{0}/{1}", subfolderPath, uploadLocation);
                }
                uploadLocation = string.Format("/{0}/{1}/{2}", webName, libraryName, uploadLocation);
                var list = clientContext.Web.Lists.GetByTitle(libraryName);
                var fileCreationInformation = new FileCreationInformation();
                fileCreationInformation.Content = System.IO.File.ReadAllBytes(fileName);
                fileCreationInformation.Overwrite = true;
                fileCreationInformation.Url = uploadLocation;
                list.RootFolder.Files.Add(fileCreationInformation);
                clientContext.ExecuteQuery();
            }
        }

        /// <summary>
        /// Downloads the file from library.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="libraryName">Name of the library.</param>
        /// <param name="subfolderPath">The subfolder path.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="downloadPath">The download path.</param>
        public void DownloadFileFromLibrary(string siteUrl, string webName, string libraryName, string subfolderPath, string fileName, string downloadPath)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                string filePath = string.Empty;
                if (!string.IsNullOrEmpty(subfolderPath))
                {
                    filePath = string.Format("/{0}/{1}/{2}/{3}", webName, libraryName, subfolderPath, fileName);
                }
                else
                {
                    filePath = string.Format("/{0}/{1}/{2}", webName, subfolderPath, fileName);
                }

                var fileInformation = File.OpenBinaryDirect(clientContext, filePath);
                var stream = fileInformation.Stream;
                IList<byte> content = new List<byte>();
                int b;
                while ((b = fileInformation.Stream.ReadByte()) != -1)
                {
                    content.Add((byte)b);
                }
                var downloadFileName = Path.Combine(downloadPath, fileName);
                System.IO.File.WriteAllBytes(downloadFileName, content.ToArray());
                fileInformation.Stream.Close();
            }
        }

        /// <summary>
        /// Deletes the file form library.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="subfolder">The subfolder.</param>
        /// <param name="attachmentFileName">Name of the attachment file.</param>
        public void DeleteFileFormLibrary(string siteUrl, string webName, string listName, string subfolder, string attachmentFileName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                string attachmentPath = string.Empty;
                if (string.IsNullOrEmpty(subfolder))
                {
                    attachmentPath = string.Format("/{0}/{1}/{2}", webName, listName, Path.GetFileName(attachmentFileName));
                }
                else
                {
                    attachmentPath = string.Format("/{0}/{1}/{2}/{3}", webName, listName, subfolder, Path.GetFileName(attachmentFileName));
                }
                var file = clientContext.Web.GetFileByServerRelativeUrl(attachmentPath);
                file.DeleteObject();
                clientContext.ExecuteQuery();
            }
        }

        /// <summary>
        /// Attaches the file to list item.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="itemId">The item id.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="overwrite">if set to <c>true</c> [overwrite].</param>
        public void AttachFileToListItem(string siteUrl, string webName, string listName, int itemId, string fileName, bool overwrite)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                FileStream fileStream = new FileStream(fileName, FileMode.Open);
                string attachmentPath = string.Format("/{0}/Lists/{1}/Attachments/{2}/{3}", webName, listName, itemId, Path.GetFileName(fileName));
                File.SaveBinaryDirect(clientContext, attachmentPath, fileStream, overwrite);
            }
        }

        /// <summary>
        /// Downloads the attached file from list item.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="itemId">The item id.</param>
        /// <param name="attachmentName">Name of the attachment.</param>
        /// <param name="listName">Name of the list.</param>
        /// <param name="downloadLocation">The download location.</param>
        public void DownloadAttachedFileFromListItem(string siteUrl, string webName, int itemId, string attachmentName, string listName, string downloadLocation)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                clientContext.Credentials = CredentialCache.DefaultNetworkCredentials;
                string attachmentPath = string.Format("/it/{0}/Lists/{1}/Attachments/{2}/{3}", webName, listName, itemId, Path.GetFileName(attachmentName));
                var fileInformation = File.OpenBinaryDirect(clientContext, attachmentPath);
                IList<byte> content = new List<byte>();
                int b;
                while ((b = fileInformation.Stream.ReadByte()) != -1)
                {
                    content.Add((byte)b);
                }
                var downloadFileName = Path.Combine(downloadLocation, attachmentName);
                System.IO.File.WriteAllBytes(downloadFileName, content.ToArray());
                fileInformation.Stream.Close();
            }
        }

        /// <summary>
        /// Deletes the attached file from list item.
        /// </summary>
        /// <param name="siteUrl">The site URL.</param>
        /// <param name="webName">Name of the web.</param>
        /// <param name="itemId">The item id.</param>
        /// <param name="attachmentFileName">Name of the attachment file.</param>
        /// <param name="listName">Name of the list.</param>
        public void DeleteAttachedFileFromListItem(string siteUrl, string webName, int itemId, string attachmentFileName, string listName)
        {
            using (ClientContext clientContext = new ClientContext(siteUrl))
            {
                //http://siteurl/lists/[listname]/attachments/[itemid]/[filename]
                string attachmentPath = string.Format("/{0}/lists/{1}/Attachments/{2}/{3}", webName, listName, itemId, Path.GetFileName(attachmentFileName));
                var file = clientContext.Web.GetFileByServerRelativeUrl(attachmentPath);
                file.DeleteObject();
                clientContext.ExecuteQuery();
            }
        }
    }
}
