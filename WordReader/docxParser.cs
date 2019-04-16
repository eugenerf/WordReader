using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace WordReader
{
    class docxParser
    {
        #region Constants
        //Path within the DOCX package to the required item rels, which contains the reference to the Main document item
        private const string relsReference = @"_rels/.rels";
        //Reference in _rels/.rels part of docx package. The target of it references to the Main Document item
        private const string docReference = @"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
        #endregion

        #region Fields
        #region private
        private XmlDocument xDocx = null;   //Main Document of the DOCX
        #endregion
        #endregion

        #region Constructors
        /// <summary>
        /// Class constructor
        /// </summary>
        /// <param name="filePath">Path to the DOCX file</param>
        protected internal docxParser(string filePath)
        {
            ZipArchive zArchive;                                            //ZIP-archive package of the specified DOCX
            ZipArchiveEntry zArchiveEntry;                                  //one entry in zArchive
            Stream zsEntry;                                                 //stream for reading zArchiveEntry
            try                                                             //trying to read specified DOCX as a ZIP-archive
            {
                zArchive = ZipFile.OpenRead(filePath);                      //open ZIP-archive
                zArchiveEntry = zArchive.GetEntry(relsReference);           //get rels item of the package
                zsEntry = zArchiveEntry.Open();                             //open zArchiveEntry as a Stream
            }
            catch (Exception)                                               //catch any exception
            {
                return;                                                     //just return without constructing the class in the case of any exception
            }
            TextReader ztrEntry = new StreamReader(zsEntry, true);          //create TextReader for zsEntry
            XmlDocument xRels = new XmlDocument();                          //create new XmlDocument to read rels
            xRels.Load(ztrEntry);                                           //load ztrEntry as XML

            XmlNode xDocNode = null;                                        //node with reference to the Main Document item of the DOCX file
            XmlNode xNode = xRels.FirstChild;                               //get first child of xRels (it'll be <?xml...> node)
            XmlNodeList xList = xNode.NextSibling.ChildNodes;               //get child nodes of the first sibling of xNode (children of the <Relationships...> node)
            foreach (XmlNode xn in xList)                                   //moving through all nodes in xList
            {
                XmlAttributeCollection xAttrCol = xn.Attributes;            //get all attributes of the current node
                foreach (XmlAttribute xa in xAttrCol)                       //moving through all the attributes of the current node
                {
                    if (xa.Name == "Type" && xa.Value == docReference)      //if current attribute is Type and its value says that current node contains the reference to the Main Document item
                    {
                        xDocNode = xAttrCol.GetNamedItem("Target");         //save attribute Target of the current node
                        break;                                              //break from the current foreach
                    }                    
                }
                if (xDocNode != null) break;                                //if node with the reference to the Main Document was found break from the current foreach
            }

            try                                                             //trying to open the MainDocument item in the specified DOCX file
            {
                zArchiveEntry = zArchive.GetEntry(xDocNode.Value);          //get the MainDocument item
                zsEntry = zArchiveEntry.Open();                             //open zArchiveEntry as a Stream
            }
            catch (Exception)                                               //catch any exception
            {
                return;                                                     //just return without constructing the class in the case of any exception
            }

            ztrEntry = new StreamReader(zsEntry, true);                     //create TextReader for zsEntry
            xDocx = new XmlDocument();                                      //create new XmlDocument to read the MainDocument
            xDocx.Load(ztrEntry);                                           //load ztrEntry to xDocx as XML
        }
        #endregion

        #region Methods
        #region private
        /// <summary>
        /// Retrieve text from the specified node
        /// </summary>
        /// <param name="xNode">Xml node</param>
        /// <param name="IsTable">If true current node is a table</param>
        /// <returns>String containing the text or empty string</returns>
        private string retrieveText(XmlNode xNode, bool IsTable = false)
        {
            string resStr = "";                                                     //result string
            if (xNode == null) return null;                                         //if got to the null node return empty string
            if (xNode.Name == "w:document" || xNode.Name == "w:body")               //if current node is document or body
                return retrieveText(xNode.FirstChild);                                  //return text just from them (we do not looking for text outside these nodes)
            do                                                                      //moving through all nodes of body
            {
                if (xNode.Name == "w:tbl")                                          //if current node is a table
                    resStr += retrieveText(xNode.FirstChild, true);                     //retrieve text from table
                if (xNode.Name == "mc:AlternateContent")                            //if current node is an AlternateContent
                {
                    xNode = xNode.NextSibling;                                          //just skip it
                    if (xNode == null) return "";                                       //if there is no more siblings return empty string
                }
                if (xNode.Name == "w:t") resStr += xNode.InnerText;                 //take inner text of the range of text node (it actually is the MainDocument text)
                if (xNode.HasChildNodes && xNode.Name != "w:tbl")                   //if current node has children and is not a table
                    resStr += retrieveText(xNode.FirstChild, IsTable);                  //retrieve text from its children
                if (xNode.Name == "w:p" && !IsTable) resStr += "\n";                //if current node is paragraph and isn't a table append \n to the result string
                xNode = xNode.NextSibling;                                          //go to the next sibling of the current node
            } while (xNode != null);                                                //do while current node is not null
            return resStr;                                                          //return result string
        }
        #endregion

        #region protected internal
        /// <summary>
        /// Get text from the DOCX-file
        /// </summary>
        /// <returns>String containing the text from DOCX-file or null</returns>
        protected internal string getText()
        {
            if (xDocx == null) return null;             //if no MainDocument has been opened, return null

            XmlNode xNode = xDocx.FirstChild;           //get the first child in xDocx (<?xml...>)
            if (xNode == null) return null;             //if there is no children within xDocx return null

            return retrieveText(xNode);     //return docxText string
        }
        #endregion
        #endregion
    }
}
