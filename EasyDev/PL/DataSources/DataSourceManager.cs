using EasyDev.Configuration;
using System.Data.Common;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据源管理器
    /// <example>
    /// 创建默认的数据源:<br/>
    /// <code>
    /// IDataSourceObject dsObj = DataSourceManager.CreateDefaultDataSource();
    /// </code>
    /// 创建命名数据源：<br/>
    /// <code>
    /// IDataSourceObject dsObj = DataSourceManager.CreateDataSource(dataSourceName);
    /// </code>
    /// <remarks>
    /// 数据源名称是在EasyDev.Persistence.Config文件中设置的名称，默认数据源是在这个配置文件中DataSources结点Default属性的值所对应的数据源配置
    /// </remarks>
    /// </example>
    /// </summary>
    public class DataSourceManager
    {
        private static PersistenceConfigManager configMgr = null;

        #region IDBProviderManager 成员

        /// <summary>
        /// 根据配置信息产生数据库提供程序
        /// </summary>
        /// <returns></returns>
        public static DataSourceObject CreateDefaultDataSource()
        {
            configMgr = PersistenceConfigManager.GetInstance();
            DataSourceObject provider = new DataSourceObject(DbProviderFactories.GetFactory(configMgr.GetDefaultDataSource().ProviderType));
            provider.ProviderName =configMgr.DefaultDataSource;
            return provider;
        }
        
        /// <summary>
        /// 根据配置信息中的数据源名称产生会话对象的方法
        /// </summary>
        /// <param name="name">数据源配置的名称</param>
        /// <returns></returns>
        public static DataSourceObject CreateDataSource(string name)
        {
            string type = "";
            configMgr = PersistenceConfigManager.GetInstance();
            foreach (IDataSource providerItem in configMgr.DataSources.Values)
            {
                if (providerItem.Name == name)
                {
                    type = providerItem.ProviderType;
                }
            }
            DataSourceObject provider = new DataSourceObject(DbProviderFactories.GetFactory(type));
            provider.ProviderName = name;
            return provider;
        }

        #endregion
    }
}
