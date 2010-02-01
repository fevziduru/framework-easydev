using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.Configuration
{
        /// <summary>
        /// 数据源对象公共接口
        /// </summary>
        public interface IDataSource
        {
                /// <summary>
                /// 数据源名称
                /// </summary>
                string Name { get; set; }

                /// <summary>
                /// 数据源提供程序类型
                /// </summary>
                string ProviderType { get; set; }

                /// <summary>
                /// 连接字符串
                /// </summary>
                string ConnectionString { get; set; }
        }
}