using System;
using System.Collections.Generic;
using System.Xml;
using System.Collections;
using System.Web;
using WebCaching = System.Web.Caching;
using EasyDev.PL;
using System.Linq;
using System.Xml.Linq;
using System.IO;
using System.Data;
using System.Security;
using EasyDev.Util;
using System.Threading;

namespace EasyDev.Configuration
{
        /// <summary>
        /// 持久层配置管理器
        /// </summary>
        public class PersistenceConfigManager
        {
                private static readonly string DEFAULT_SCHEMA_DIR = "DEFAULT_SCHEMA_DIR";

                private static readonly string EasyDev_DATASOURCE_CONFIG = "EasyDev_DATASOURCE_CONFIG";

                private static readonly string DEFAULT_DATASOURCE = "DEFAULT_DATASOURCE";

                private FileSystemWatcher watcher = null;

                /// <summary>
                /// 默认数据提供程序
                /// </summary>
                private string defaultDataSource = string.Empty;

                /// <summary>
                /// 默认数据架构目录
                /// </summary>
                private string defaultSchemaDir = string.Empty;

                /// <summary>
                /// 数据源缓存
                /// </summary>
                private WebCaching.Cache DataSourceCache { get; set; }

                /// <summary>
                /// 配置文件位置
                /// </summary>
                private string configPath = string.Empty;

                /// <summary>
                /// 数据提供程序集合
                /// </summary>
                private IDictionary<string, IDataSource> datasources = null;

                /// <summary>
                /// 锁对象
                /// </summary>
                private static readonly object lockHelper = new object();

                /// <summary>
                /// 持久配置对象管理器实例
                /// </summary>
                private static PersistenceConfigManager _instance = null;

                private bool IsConfigUpdated = false;

                /// <summary>
                /// 默认数据提供程序
                /// </summary>
                public string DefaultDataSource
                {
                        get { return this.defaultDataSource; }
                        set { this.defaultDataSource = value; }
                }

                /// <summary>
                /// 默认数据架构目录
                /// </summary>
                public string SchemaDir
                {
                        get { return this.defaultSchemaDir; }
                        set { this.defaultSchemaDir = value; }
                }

                /// <summary>
                /// 数据提供程序集合
                /// </summary>
                public IDictionary<string, IDataSource> DataSources
                {
                        get { return this.datasources; }
                }

                private PersistenceConfigManager()
                {
                        this.configPath = AppDomain.CurrentDomain.BaseDirectory + @"/Config/EasyDev.Persistence.Config";
                        this.datasources = new Dictionary<string, IDataSource>();
                        watcher = new FileSystemWatcher();
                        watcher.Filter = "EasyDev.Persistence.Config";
                        watcher.Path = AppDomain.CurrentDomain.BaseDirectory + @"Config\\";
                        watcher.Renamed += new RenamedEventHandler(watcher_Renamed);
                        watcher.Changed += new FileSystemEventHandler(watcher_Changed);
                        watcher.Deleted += new FileSystemEventHandler(watcher_Deleted);
                        watcher.EnableRaisingEvents = true;
                        watcher.IncludeSubdirectories = false;
                        this.Initialize();
                }

                void watcher_Deleted(object sender, FileSystemEventArgs e)
                {
                        HttpRuntime.Cache.Remove(EasyDev_DATASOURCE_CONFIG);
                        HttpRuntime.Cache.Remove(DEFAULT_DATASOURCE);
                        HttpRuntime.Cache.Remove(DEFAULT_SCHEMA_DIR);
                        this.datasources.Clear();
                }

                void watcher_Changed(object sender, FileSystemEventArgs e)
                {
                        //数据源配置文件的位置固定
                        HttpRuntime.Cache.Remove(EasyDev_DATASOURCE_CONFIG);
                        HttpRuntime.Cache.Remove(DEFAULT_DATASOURCE);
                        HttpRuntime.Cache.Remove(DEFAULT_SCHEMA_DIR);
                        FetchPersistenceConfig();
                }

                void watcher_Renamed(object sender, RenamedEventArgs e)
                {
                        //数据源配置文件的位置固定
                        HttpRuntime.Cache.Remove(EasyDev_DATASOURCE_CONFIG);
                        HttpRuntime.Cache.Remove(DEFAULT_DATASOURCE);
                        HttpRuntime.Cache.Remove(DEFAULT_SCHEMA_DIR);
                        FetchPersistenceConfig();
                }

                /// <summary>
                /// 创建持久层对象实例
                /// </summary>
                /// <returns></returns>
                public static PersistenceConfigManager GetInstance()
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new PersistenceConfigManager();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 初始化数据提供程序配置
                /// </summary>
                private void Initialize()
                {
                        FetchPersistenceConfig();
                }

                private void FetchPersistenceConfig()
                {
                        this.datasources = HttpRuntime.Cache.Get(EasyDev_DATASOURCE_CONFIG) as Dictionary<string, IDataSource>;


                        if (this.datasources == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(HttpRuntime.Cache, 100))
                                        {
                                                if (this.datasources == null)
                                                {
                                                        InitDataSourceConfig();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(HttpRuntime.Cache);
                                }
                        }
                }

                /// <summary>
                /// 初始化数据源
                /// </summary>
                [Obsolete("禁用此方法", true)]
                private void InitDataSource()
                {
                        IEnumerable<XElement> nodes =
                                from p in XElement.Load(AppDomain.CurrentDomain.BaseDirectory + @"Config\\EasyDev.Persistence.Config").Elements("DataSources") select p;

                        if (nodes != null && nodes.Count<XElement>() > 0)
                        {
                                XElement rootNode = nodes.First<XElement>();

                                if (rootNode.Attribute("SchemaDir") != null)
                                {
                                        this.defaultSchemaDir = AppDomain.CurrentDomain.BaseDirectory + @"\" + rootNode.Attribute("SchemaDir").Value;
                                        HttpRuntime.Cache.Insert(DEFAULT_SCHEMA_DIR, this.defaultSchemaDir);
                                }

                                if (rootNode.Attribute("Default") != null)
                                {
                                        this.defaultDataSource = rootNode.Attribute("Default").Value;
                                        HttpRuntime.Cache.Insert(DEFAULT_DATASOURCE, this.defaultDataSource);
                                }

                                IEnumerable<XElement> datasources = rootNode.Elements("DataSource");
                                foreach (XElement item in datasources)
                                {
                                        IDataSource newDatasource = new DataSource();
                                        XElement nameNode = item.Elements().First<XElement>(p => p.Name.LocalName.Equals("Name"));
                                        XElement providerNode = item.Elements().First<XElement>(p => p.Name.LocalName.Equals("ProviderType"));
                                        XElement connstrNode = item.Elements().First<XElement>(p => p.Name.LocalName.Equals("ConnectionString"));

                                        if (nameNode.Value.Length == 0)
                                        {
                                                throw new System.Exception("_datasource_must_have_a_name");
                                        }

                                        newDatasource.Name = nameNode.Value;
                                        newDatasource.ProviderType = providerNode.Value;
                                        newDatasource.ConnectionString = connstrNode.Value;

                                        KeyValuePair<string, IDataSource> ds = new KeyValuePair<string, IDataSource>(nameNode.Value, newDatasource);
                                        if (this.datasources == null)
                                        {
                                                this.datasources = new Dictionary<string, IDataSource>();
                                        }

                                        if (this.datasources.Keys.Contains(ds.Key) == false)
                                        {
                                                this.datasources.Add(ds);
                                        }
                                }

                                HttpRuntime.Cache.Insert(EasyDev_DATASOURCE_CONFIG, this.datasources);
                        }
                        else
                        {
                                throw new PersistenceException("_persistence_config_lack_of_root");
                        }
                }

                private void InitDataSourceConfig()
                {
                        DataSet dsConfig = new DataSet();

                        try
                        {
                                Thread.Sleep(500);
                                using (FileStream fstream =
                                        new FileStream(AppDomain.CurrentDomain.BaseDirectory + @"Config\\EasyDev.Persistence.Config", FileMode.Open, FileAccess.Read))
                                {
                                        dsConfig.ReadXml(fstream);
                                }
                        }
                        catch (SecurityException e)
                        {
                                throw e;
                        }

                        if (dsConfig.Tables.Count > 0)
                        {
                                DataRow drDataSourcesConfig = dsConfig.FirstRow("DataSources");
                                DataTable dtDataSourceItems = dsConfig.Table("DataSource");

                                this.defaultSchemaDir = AppDomain.CurrentDomain.BaseDirectory + @"\" + drDataSourcesConfig.Field<string>("SchemaDir");
                                HttpRuntime.Cache.Insert(DEFAULT_SCHEMA_DIR, this.defaultSchemaDir);

                                this.defaultDataSource = drDataSourcesConfig.Field<string>("Default");
                                HttpRuntime.Cache.Insert(DEFAULT_DATASOURCE, this.defaultDataSource);

                                foreach (DataRow row in dtDataSourceItems.Rows)
                                {
                                        IDataSource newDatasource = new DataSource();
                                        if (row.Field<string>("Name").Length == 0)
                                        {
                                                throw new System.Exception("_datasource_must_have_a_name");
                                        }

                                        newDatasource.Name = row.Field<string>("Name");
                                        newDatasource.ProviderType = row.Field<string>("ProviderType");
                                        newDatasource.ConnectionString = row.Field<string>("ConnectionString");

                                        KeyValuePair<string, IDataSource> ds = new KeyValuePair<string, IDataSource>(newDatasource.Name, newDatasource);
                                        if (this.datasources == null)
                                        {
                                                this.datasources = new Dictionary<string, IDataSource>();
                                        }

                                        if (this.datasources.Keys.Contains(ds.Key) == false)
                                        {
                                                this.datasources.Add(ds);
                                        }
                                }

                                HttpRuntime.Cache.Insert(EasyDev_DATASOURCE_CONFIG, this.datasources);
                        }
                        else
                        {
                                throw new PersistenceException("_persistence_config_lack_of_root");
                        }
                }

                /// <summary>
                /// 取得默认数据提供程序信息
                /// </summary>
                /// <returns></returns>
                public IDataSource GetDefaultDataSource()
                {
                        IDictionary<string, IDataSource> ds = HttpRuntime.Cache.Get(EasyDev_DATASOURCE_CONFIG) as Dictionary<string, IDataSource>;


                        if (ds == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(HttpRuntime.Cache, 100))
                                        {
                                                if (ds == null)
                                                {
                                                        InitDataSourceConfig();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(HttpRuntime.Cache);
                                }
                        }

                        return this.datasources[this.defaultDataSource];
                }

                /// <summary>
                /// 根据名称取得数据提供程序信息
                /// </summary>
                /// <param name="name"></param>
                /// <returns></returns>
                public IDataSource GetDataSourceByName(string name)
                {
                        IDictionary<string, IDataSource> ds = HttpRuntime.Cache.Get(EasyDev_DATASOURCE_CONFIG) as Dictionary<string, IDataSource>;


                        if (ds == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(HttpRuntime.Cache, 100))
                                        {
                                                if (ds == null)
                                                {
                                                        InitDataSourceConfig();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(HttpRuntime.Cache);
                                }
                        }

                        return this.datasources[name];
                }
        }
}