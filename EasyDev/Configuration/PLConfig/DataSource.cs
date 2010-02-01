using System;
using System.Data;

namespace EasyDev.Configuration
{
        /// <summary>
        /// 数据源对象
        /// </summary>
        public sealed class DataSource : IDataSource
        {
                /// <summary>
                /// 数据源名称
                /// </summary>
                public string Name
                {
                        get;
                        set;
                }

                /// <summary>
                /// 提供程序类型
                /// </summary>
                public string ProviderType
                {
                        get;
                        set;
                }

                /// <summary>
                /// 连接字符串
                /// </summary>
                public string ConnectionString
                {
                        get;
                        set;
                }
        }
}