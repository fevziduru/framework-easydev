using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Collections;
using EasyDev.Resources;

namespace EasyDev.Resources
{
    public abstract class GlobalizerBase : IGlobalizer
    {
        protected XmlDocument languageNodes = null;
        
        public GlobalizerBase(XmlDocument languageInfo)
        {
            this.languageNodes = languageInfo;
        }

        /// <summary>
        /// 在国际化语言信息中查找对应的值
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        protected string GetValueByKey(string key)
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
        
        #region IGlobalizer 成员

        public virtual string GetString(string key)
        {
            throw new System.NotSupportedException("The method or operation is not supported.");
        }

        public virtual string GetFormatString(string key, params object[] objs)
        {
            throw new System.NotSupportedException("The method or operation is not supported.");
        }

        #endregion
    }
}
