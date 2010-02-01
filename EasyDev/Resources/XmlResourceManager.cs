using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Configuration;
using System.Xml;
using System.Collections;
using System.IO;
using System.Threading;
using System.Reflection;

namespace EasyDev.Resources
{
    public class XmlResourceManager : ResourceManagerBase
    {
        private XmlDocument languageNodes = null;

        public XmlResourceManager() { }

        #region ResourceManagerBase Members

        public override string GetString(string key)
        {
            return this.GetValueByKey(key);
        }

        public override void Initialize(ResourceItemConfig item)
        {
            string cultureName = item.Culture;
            if (string.IsNullOrEmpty(cultureName))
            {
                cultureName = Thread.CurrentThread.CurrentCulture.Name;
            }
            string path = string.Format(@"{0}.{1}_{2}.xml", item.Assembly, item.BaseName, cultureName);
            Stream resouceStream = Assembly.Load(item.Assembly).GetManifestResourceStream(path);
            StreamReader reader = null;

            try
            {
                if (resouceStream != null && resouceStream.CanRead)
                {
                    resouceStream.Seek(0, SeekOrigin.Begin);
                    reader = new StreamReader(resouceStream);
                    this.languageNodes = new XmlDocument();
                    this.languageNodes.LoadXml(reader.ReadToEnd());
                }
                else
                {
                    throw new GlobalizeException("_cannot_find_language_resource");
                }
            }
            catch (GlobalizeException e)
            {
                throw e;
            }
        }

        #endregion

        /// <summary>
        /// 在国际化语言信息中查找对应的值
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        private string GetValueByKey(string key)
        {
            string value = string.Empty;
            try
            {
                XmlNodeList items = this.languageNodes.SelectNodes("./Resources/ResourceItem");
                XmlNode item = null;
                XmlNode node = null;
                IEnumerator itr_items = items.GetEnumerator();
                while (itr_items.MoveNext())
                {
                    node = ((XmlNode)itr_items.Current);
                    item = node.SelectSingleNode("Key");

                    if (item != null && item.InnerText.Equals(key))
                    {
                        break;
                    }
                }

                if (node != null && node.HasChildNodes)
                {
                    value = node.SelectSingleNode("Value").InnerText;
                }
                else
                {
                    throw new GlobalizeException("_can_not_find_the_right_language_item");
                }
            }
            catch (GlobalizeException e)
            {
                throw e;
            }

            return value;
        }
    }
}
