using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Configuration;
using System.Resources;
using System.Reflection;
using System.Threading;
using System.Globalization;
using System.Collections;

namespace EasyDev.Resources
{
    public class StandardResourceManager : ResourceManagerBase
    {        
        private ResourceSet resources = null;

        public StandardResourceManager() 
        {

        }

        #region ResourceManagerBase Members

        /// <summary>
        /// 取得资源串值
        /// </summary>
        /// <param name="key">资源的唯一标识</param>
        /// <returns></returns>
        public override string GetString(string key)
        {
            return this.resources.GetString(key);
        }

        public override void Initialize(ResourceItemConfig item)
        {
            ResourceManager manager = new ResourceManager(item.BaseName, Assembly.Load(item.Assembly));
            string cultureName = item.Culture;

            //如果没有设置文化区域名称则使用当前线程的设置
            if (string.IsNullOrEmpty(cultureName))
            {
                cultureName = Thread.CurrentThread.CurrentCulture.Name;
            }

            this.resources = manager.GetResourceSet(new CultureInfo(cultureName), true, true);
        }

        #endregion

        public object GetObject(string key)
        {
            return GetObject(key, false);
        }

        public object GetObject(string key, bool isIgnoreCase)
        {
            return this.resources.GetObject(key, isIgnoreCase);
        }

        public IEnumerable GetEnumeration()
        {
            return this.GetEnumeration();
        }        
    }
}
