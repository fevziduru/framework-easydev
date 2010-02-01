using System;
using System.Reflection;
using System.Configuration;
using System.Runtime.Remoting;
using EasyDev.PL;
using EasyDev.Util;
using System.Threading;

namespace EasyDev.DAL
{
        public enum AssemblyNameSource
        {
                FromString,
                FromConfiguration
        }

        /// <summary>
        /// 数据库连接管理器
        /// </summary>
        public class GenericDBFactory
        {
                /// <summary>
                /// 同步锁对象
                /// </summary>
                private static object _lockHelper = new object();

                /// <summary>
                /// 管理器实例
                /// </summary>
                private static GenericDBFactory _instance = null;

                #region 公共属性

                //public MSSqlSession MSSqlSession
                //{
                //    get
                //    {
                //        return this.CreateDBSession<MSSqlSession>();
                //    }
                //}

                //public OracleSession OracleSession
                //{
                //    get
                //    {
                //        return this.CreateDBSession<OracleSession>();
                //    }
                //}

                //public AccessSession AccessSession
                //{
                //    get
                //    {
                //        return this.CreateDBSession<AccessSession>();
                //    }
                //}

                #endregion

                protected GenericDBFactory()
                { }

                /// <summary>
                /// 取得管理器实例
                /// </summary>
                public static GenericDBFactory GetInstance()
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new GenericDBFactory();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }
                        return _instance;
                }

                /// <summary>
                /// (泛型方法) 创建DBSession对象，这个对象必须实现IDBSession接口，并且Session对象必须要有无参数的公共构造函数
                /// </summary>
                /// <typeparam name="T">泛型类型</typeparam>
                /// <param name="dataSourceName">位于配置文件中的数据源名称</param>
                /// <returns>对应泛型类型的实例化对象</returns>
                public T CreateDBSession<T>()
                    where T : IGenericDBSession, new()
                {
                        return new T();
                }

                /// <summary>
                /// 根据程序集名称动态加载程序集并且实例化数据访问对象
                /// 数据集中的数据访问对象必须实现IDBSession接口
                /// </summary>
                /// <typeparam name="T">泛型类型</typeparam>
                /// <param name="name">程序集名称</param>
                /// <param name="source">程序集名称类型(名称字符串，配置字符串)</param>
                /// <returns>对应泛型类型的实例化对象</returns>
                public T CreateDBSession<T>(string name, AssemblyNameSource source)
                    where T : IGenericDBSession, new()
                {
                        Assembly assembly = null;

                        try
                        {
                                string assemblyName = "";

                                if (source == AssemblyNameSource.FromConfiguration)
                                {
                                        assemblyName = ConvertUtil.ConvertToString(ConfigurationManager.AppSettings[name]);
                                }
                                else if (source == AssemblyNameSource.FromString)
                                {
                                        assemblyName = name;
                                }
                                else
                                {
                                        throw new DALCoreException("assembly_name_is_not_allowed"); //程序集名称来源不合法
                                }

                                //根据程序集名称创建程序集引用
                                assembly = Assembly.Load(new AssemblyName(assemblyName));
                                T component = Activator.CreateInstance<T>();

                                //判断外部数据访问组件是否继承于DBSessionBase，如果不是则抛出异常
                                if ((component is IGenericDBSession) == false)
                                {
                                        throw new DALCoreException("DAComponent_must_be_derived_from_DBSessionBase");    //数据访问组件必须继承于DBSessionBase
                                }

                                return component;
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 根据类型串生成数据库会话对象
                /// </summary>
                /// <param name="type">类型串</param>
                /// <returns>返回IDBSession类型对象</returns>
                public IGenericDBSession CreateDBSession(string type)
                {
                        ObjectHandle handle = null;
                        try
                        {
                                if (string.IsNullOrEmpty(type) || type.IndexOf(',') == -1)
                                {
                                        throw new DALCoreException("_type_is_not_vaild");
                                }
                                string[] assemblyInfo = type.Split(',');
                                handle = Activator.CreateInstance(assemblyInfo[1], assemblyInfo[0]);
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }

                        return (IGenericDBSession)handle.Unwrap();
                }

                /// <summary>
                /// 利用反射创建数据访问对象
                /// </summary>
                /// <typeparam name="T"></typeparam>
                /// <param name="assemblyName">程序集名称</param>
                /// <param name="typeName">类型名称</param>
                /// <returns></returns>
                public T CreateDBSession<T>(string assemblyName, string typeName) where T : IGenericDBSession
                {
                        try
                        {
                                ObjectHandle handle = Activator.CreateInstance(assemblyName, typeName);

                                T component = (T)handle.Unwrap();

                                if ((component is IGenericDBSession) == false)
                                {
                                        throw new DALCoreException("DAComponent_must_be_derived_from_DBSessionBase");    //数据访问组件必须继承于DBSessionBase
                                }

                                return component;
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 根据Web.Config中的数据连接字符串创建数据库会话对象
                /// </summary>
                /// <typeparam name="T"></typeparam>
                /// <param name="cfgName">配置于web.config文件中AppSetting节的数据库连接串</param>
                /// <returns></returns>
                public T CreateDBSessionByConfig<T>(string cfgName) where T : IGenericDBSession
                {
                        string[] cfg = null;
                        string assemblyName = "";
                        string typeName = "";

                        try
                        {
                                if (string.IsNullOrEmpty(ConvertUtil.ConvertToString(ConfigurationManager.AppSettings[cfgName])) == false)
                                {
                                        cfg = ConvertUtil.ConvertToString(ConfigurationManager.AppSettings[cfgName]).Split(new char[] { ',' });
                                }
                                else
                                {
                                        throw new DALCoreException("_there_is_not_configuration_for_session");
                                }

                                if (cfg != null && cfg.Length >= 2)
                                {
                                        assemblyName = cfg[0];
                                        typeName = cfg[1];
                                }
                                else
                                {
                                        throw new ArgumentException("_there_is_not_enough_configuration", "in_method_CreateDBSession<T>(string cfgName)");
                                }

                                return this.CreateDBSession<T>(assemblyName, typeName);
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }
                }
        }
}
