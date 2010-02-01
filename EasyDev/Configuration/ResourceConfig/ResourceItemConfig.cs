using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.Configuration
{
    /// <summary>
    /// 资源配置项
    /// </summary>
    public class ResourceItemConfig
    {
        /// <summary>
        /// 资源名称
        /// </summary>
        public string Name
        {
            get;
            set;
        }

        /// <summary>
        /// 资源所在程序集
        /// </summary>
        public string Assembly
        {
            get;
            set;
        }

        /// <summary>
        /// 资源基名称(不包含区域信息的名称部分)
        /// </summary>
        public string BaseName
        {
            get;
            set;
        }

        /// <summary>
        /// 文化区域名称
        /// </summary>
        public string Culture
        {
            get;
            set;
        }
    }
}
