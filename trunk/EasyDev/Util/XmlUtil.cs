using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace EasyDev.Util
{
    public class XmlUtil
    {   
        private static XmlDocument GetXmlDocument(string path)
        {
            XmlDocument doc = null;
            try
            {
                doc = new XmlDocument();
                doc.Load(path);
            }
            catch (System.Exception e)
            {
                throw e;
            }

            return doc;
        }

        public static XmlNode FindNodeByID(string path, string id)
        {
            XmlDocument doc = GetXmlDocument(path);
            XmlNode node = null;
            try
            {
                node = doc.GetElementById(id);
            }
            catch (System.Exception e)
            {
                throw e;
            }

            return node;
        }

        public static XmlNodeList FindNodesByTagName(string path, string tagName)
        {
            XmlDocument doc = GetXmlDocument(path);
            XmlNodeList list = null;

            try
            {
                list = doc.GetElementsByTagName(tagName);
            }
            catch (System.Exception e)
            {
                throw e;
            }

            return list;
        }

        public static XmlNode FindNodeByAttribute(string path, string attributeName, string attributeValue)
        {
            XmlDocument doc = GetXmlDocument(path);
            XmlNode node = null;
            
            try
            {
                XmlNodeList nodeList = doc.ChildNodes;
                for (int i = 0; i < nodeList.Count; i++)
                {
                    if (nodeList[i].Attributes[attributeName].ToString().Equals(attributeValue))
                    {
                        node = nodeList[i];
                        break;
                    }
                }
            }
            catch (System.Exception e)
            {
                throw e;
            }
            
            return node;
        }
    }
}
