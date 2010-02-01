using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Configuration;
using System.Threading;

namespace EasyDev.Resources
{
        /// <summary>
        /// 资源管理器工厂,用于创建不同类型的资源管理器
        /// </summary>
        public class ResourceManagerFactory
        {
                private ResourcesConfigManager resourceMgr = null;
                private static readonly object _lockHelper = new object();
                private static ResourceManagerFactory _instance = null;

                private ResourceManagerFactory()
                {
                        resourceMgr = ResourcesConfigManager.GetInstance();
                }

                /// <summary>
                /// 创建唯一实例
                /// </summary>
                /// <returns></returns>
                public static ResourceManagerFactory GetInstance()
                {
                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new ResourceManagerFactory();
                                                }
                                        }
                                }
                                catch (ArgumentException arge)
                                {
                                        throw arge;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 创建资源管理器
                /// </summary>
                /// <typeparam name="T">资源管理器的类型</typeparam>
                /// <param name="name">资源配置名称</param>
                /// <returns></returns>
                [Obsolete("此方法在没有配置信息的情况下不能正确产生资源管理器,使用'T CreateResourceManager<T,K>(string name)'可以在没有配置信息的情况下直接读取资源文件")]
                public T CreateResourceManager<T>(string name) where T : ResourceManagerBase, new()
                {
                        ResourceItemConfig item = null;
                        if (this.resourceMgr.IsResourceItemExist(name))
                        {
                                item = this.resourceMgr.GetResourceItemByName(name);
                        }

                        T rm = new T();
                        rm.Initialize(item);
                        return rm;
                }

                /// <summary>
                /// 创建资源管理器
                /// 此方法会先到配置文件中查找有没有对应名称的配置项，如果没有则根据类型信息自动创建资源管理器
                /// </summary>
                /// <typeparam name="T">资源管理器的类型</typeparam>
                /// <param name="name">资源配置名称，如果在配置中找不到资源配置信息则直接创建资源管理器，直接创建的资源管理器
                /// 直接创建的资源管理器，其基名称为K类型所在的命名空间名称加类型名称，程序集为配置文件中Resources结点的DefaultAssembly属性值,文化区域由当前线程的设置决定</param>
                /// <returns></returns>
                public T CreateResourceManager<T, K>(string name) where T : ResourceManagerBase, new()
                {
                        ResourceItemConfig item = null;
                        if (this.resourceMgr.IsResourceItemExist(name))
                        {
                                item = this.resourceMgr.GetResourceItemByName(name);
                        }
                        else
                        {
                                item = new ResourceItemConfig();
                                item.Assembly = this.resourceMgr.DefaultResrouceAssembly;
                                item.BaseName = this.resourceMgr.DefaultResrouceAssembly + "." + typeof(K).Namespace + "." + typeof(K).Name;
                                item.Culture = Thread.CurrentThread.CurrentCulture.Name;
                        }

                        T rm = new T();
                        rm.Initialize(item);
                        return rm;
                }

                /// <summary>
                /// 创建资源管理器
                /// </summary>
                /// <typeparam name="T">资源管理器的类型</typeparam>
                /// <typeparam name="K">使用资源的对象类型,以此类型的类型名称（不包含命名空间）作为资源配置名称</typeparam>
                /// <returns></returns>
                public T CreateResourceManager<T, K>() where T : ResourceManagerBase, new()
                {
                        return CreateResourceManager<T, K>(typeof(K).ToString());
                }
        }
}
