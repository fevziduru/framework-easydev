using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Configuration;

namespace EasyDev.Resources
{
    public abstract class ResourceManagerBase : IResourceManager
    {
        #region IResourceManager Members

        public virtual string GetString(string key)
        {
            return "None";
        }

        #endregion
        /// <summary>
        /// 初始化资源管理器
        /// </summary>
        public abstract void Initialize(ResourceItemConfig item);
    }
}
