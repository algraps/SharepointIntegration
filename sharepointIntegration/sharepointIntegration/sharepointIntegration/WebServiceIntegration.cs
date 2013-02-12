/*
 * Use webService to integrate sharepoint
 * Author: Alessandro Graps
 * Year: 2013
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Xml;

namespace sharepointIntegration
{
    public static class WebServiceIntegration
    {
        private static ICredentials credentials = CredentialCache.DefaultCredentials; //Default Credentials
        private static string ListId = "{221A4262-5B61-4B93-B7BC-50C7247EEBC9}"; //Id of the custom list thai i use
        
        /// <summary>
        /// Uploads the file.
        /// </summary>
        /// <param name="destinationPath">The destination path.</param>
        /// <param name="destinationFolderPath">The destination folder path.</param>
        /// <param name="sourceFilePath">The source file path.</param>
        /// <returns></returns>
        public static int UploadFile(string destinationPath, string destinationFolderPath, string sourceFilePath)
        {
            /*
            // Example of parameters for this function
 
            //Parameter 1
            string destinationPath = <SiteUrl> + "/" + <DocumentLibraryName> + "/" + <FolderName1> 
                                       + "/" + <FolderName2> + "/" + <FileName>;
            //Example
            string destinationPath = "http://sharepoint.site.com" + "/" + "HistoricalDocs" 
                                        + "/" + "Historical" + "/" + "IT" + "/" + "Historical_IT.doc";
 
            //Parameter 2
            string destinationFolderPath = <SiteUrl> + "/" + <DocumentLibraryName> + "/" + <FolderName1> 
                                             + "/" + <FolderName2>;
            //Example
            string destinationFolderPath = "http://sharepoint.site.com" + "/" + "HistoricalDocs"  
                                             + "/" + "Historical" + "/" + "IT";
 
            //Parameter 3
            string sourceFilePath = "c:\HistoricalDocs\Historical\20121111\IT\Historical_IT.doc;
            */

            int result = -1; // -1 means failure
            Copy.Copy copyReference = new Copy.Copy();
            try
            {
                string destination = destinationPath;
                string[] destinationUrl = { destination };
                byte[] fileBytes = GetFileFromFileSystem(sourceFilePath);
                copyReference.Url = "http://sharepoint.site.com" + "/_vti_bin/copy.asmx";
                
                //copyReference.Credentials = new NetworkCredential(settings.User, settings.Password, settings.Domain); ;
                copyReference.Credentials = CredentialCache.DefaultCredentials;

                Copy.CopyResult[] cResultArray = null;
                Copy.FieldInformation documentTypeField = new Copy.FieldInformation();
                documentTypeField.DisplayName = "DocumentType";
                documentTypeField.Type = Copy.FieldType.Text;
                documentTypeField.Value = "HR";
                Copy.FieldInformation titleField = new Copy.FieldInformation();
                titleField.DisplayName = "Title";
                titleField.Type = Copy.FieldType.Text;
                titleField.Value = "HR1";

                //Associate metadata
                Copy.FieldInformation[] filedInfo = { documentTypeField, titleField };
                //Upload the document from Local to SharePoint
                copyReference.CopyIntoItems(destination, destinationUrl, filedInfo, fileBytes, out cResultArray);
                //Get item id of uploaded docuemnt // fileName will be like abc.pdf
                string fileName = Path.GetFileName(destinationPath);
                string queryXml = "<Where><And><Eq><FieldRef Name=’FileLeafRef’ /><Value Type=’File’>" + fileName +
                 "</Value></Eq><Eq><FieldRef Name=’ContentType’ /><Value Type=’Text’>Document</Value></Eq></And></Where>";
                string itemId = GetItemId(ListId, queryXml, destinationFolderPath);

                result = Convert.ToInt32(itemId);
            }
            catch (Exception ex)
            {//exception code here 
            }
            finally
            {
                if (copyReference != null)
                    copyReference.Dispose();
            }
            return result;
        }

        /// <summary>
        /// Gets the file from file system.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <returns></returns>
        private static byte[] GetFileFromFileSystem(string path)
        {
            byte[] fileBytes = null;

            if (File.Exists(path))
            {
                //read the file.
                using (FileStream fs = File.Open(path, FileMode.Open))
                {
                    fileBytes = new byte[fs.Length];
                    fs.Position = 0;
                    fs.Read(fileBytes, 0, Convert.ToInt32(fs.Length));
                }
            }
            return fileBytes;
        }

        /// <summary>
        /// Gets the item ID.
        /// </summary>
        /// <param name="strListName">Name of the STR list.</param>
        /// <param name="queryXml">The query XML.</param>
        /// <param name="destinationFolderPath">The destination folder path.</param>
        /// <returns></returns>
        public static string GetItemId(string strListName, string queryXml, string destinationFolderPath)
        {
	        Lists.Lists listReference = new Lists.Lists();
	        string lookupId = string.Empty;
	        try
	        {
		        listReference.Credentials = credentials;
		        listReference.Url = "http://sharepoint.site.com"  + "/_vti_bin/lists.asmx";
		        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
		        System.Xml.XmlElement query = xmlDoc.CreateElement("Query");
		        System.Xml.XmlElement viewFields = xmlDoc.CreateElement("ViewFields");
		        System.Xml.XmlElement queryOptions = xmlDoc.CreateElement("QueryOptions");
		        query.InnerXml = queryXml;
		        viewFields.InnerXml = "<FieldRef Name=\"ID\" />";
		        queryOptions.InnerXml = "<Folder>" + destinationFolderPath + "</Folder>";
		        System.Xml.XmlNode items;
		        if (String.IsNullOrEmpty(destinationFolderPath))
			        items = listReference.GetListItems(strListName, string.Empty, query, viewFields, string.Empty, null, null);
		        else
			        items = listReference.GetListItems(strListName, string.Empty, query, viewFields, string.Empty, 
                                                                   queryOptions, null);
 
		        foreach (System.Xml.XmlNode node in items)
		        {
			        if (node.Name == "rs:data")
			        {
				        for (int i = 0; i < node.ChildNodes.Count; i++)
				        {
					        if (node.ChildNodes[i].Name == "z:row")
					        {
						        lookupId = node.ChildNodes[i].Attributes["ows_ID"].Value;
						        break;
					        }
				        }
			        }
		        }
	        }
	        catch (Exception ex)
	        {
		        //exception
	        }
	        finally
	        {
		        if (listReference != null)
			        listReference.Dispose();
	        }
 
	        return lookupId;
        }


        /// <summary>
        /// Updates the meta data.
        /// </summary>
        /// <param name="strListName">Name of the STR list.</param>
        /// <param name="strCaml">The STR CAML.</param>
        public static void UpdateMetaData(string strListName, string strCaml)
        {
            Lists.Lists listReference = new Lists.Lists();
            try
            {
                listReference.Credentials = credentials;
                listReference.Url = "http://sharepoint.site.com" + "/_vti_bin/lists.asmx";

                /*Get Name attribute values (GUIDs) for list and view. */
                System.Xml.XmlNode ndListView = listReference.GetListAndView(strListName, "");
                string strListID = ndListView.ChildNodes[0].Attributes["Name"].Value;
                string strViewID = ndListView.ChildNodes[1].Attributes["Name"].Value;

                /*Create an XmlDocument object and construct a Batch element and its
                 attributes. Note that an empty ViewName parameter causes the method to use the default view. */
                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                System.Xml.XmlElement batchElement = doc.CreateElement("Batch");
                batchElement.SetAttribute("OnError", "Continue");
                batchElement.SetAttribute("ListVersion", "1");
                batchElement.SetAttribute("ViewName", strViewID);

                /*Specify methods for the batch post using CAML. To update or delete, 
                 specify the ID of the item, and to update or add, specify 
                 the value to place in the specified column.*/
                batchElement.InnerXml = strCaml;

                /*Update list items. This example uses the list GUID, which is recommended, 
                 but the list display name will also work.*/
                listReference.UpdateListItems(strListID, batchElement);
            }
            catch (Exception ex)
            {
                //exception
            }
            finally
            {
                if (listReference != null)
                    listReference.Dispose();
            }
        }

        /// <summary>
        /// Createcontents the type.
        /// </summary>
        /// <returns></returns>
        public static string CreatecontentType()
        {
            Lists.Lists listReference = new Lists.Lists();
            string lookupId = string.Empty;
            try
            {
                XmlDocument xmlDoc = new XmlDocument();

                XmlNode xmlProps = xmlDoc.CreateNode(XmlNodeType.Element, "ContentType", "");

                XmlAttribute xmlPropsDesc = xmlDoc.CreateAttribute("Description");
                xmlPropsDesc.Value = "New content type description";
                xmlProps.Attributes.Append(xmlPropsDesc);

                listReference.Credentials = CredentialCache.DefaultNetworkCredentials;
                listReference.Url = "http://sharepoint.site.com" + "/_vti_bin/lists.asmx";


                listReference.CreateContentType(ListId, "NewContentType", "0x0101",
                   xmlDoc.CreateNode(XmlNodeType.Element, "FieldRefs", ""), xmlProps, "true");

            }
            catch (Exception ex)
            {
                //exception
            }
            finally
            {
                if (listReference != null)
                    listReference.Dispose();
            }

            return lookupId;
        }

    }

}
