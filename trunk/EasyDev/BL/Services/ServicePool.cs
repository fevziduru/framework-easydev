using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Reflection;
using System.Web;

namespace EasyDev.BL.Services
{
    /// <summary>
    /// 服务池对象
    /// </summary>
    public class ServicePool
    {
        /// <summary>
        /// 从指定的程序集将服务加载到缓存
        /// </summary>
        public static void LoadServices()
        {
            IDictionary<string, IService> services = new Dictionary<string, IService>();
            string[] assemblyName = 
                ConfigurationManager.AppSettings["ServiceGallery"].Split(new char[] { ',' });
            NativeServiceManager ServiceManager = ServiceManagerFactory.CreateNativeServiceManager();

            foreach (string item in assemblyName)
            {
                Assembly ass = Assembly.Load(new AssemblyName(item));
                Type[] types = ass.GetTypes();
                foreach (Type type in types)
                {
                    if (type.GetInterface("IService") != null)
                    {
                        services.Add(type.FullName, ServiceManager.CreateService(type));
                    }
                }
            }

            HttpRuntime.Cache.Insert("SERVICE_GALLERY", services);
        }

        /// <summary>
        /// 从服务池中取得已经实例化的服务
        /// </summary>
        /// <param name="key">服务的完全限定名</param>
        /// <returns></returns>
        public static IService FetchService(string key)
        {
            IDictionary<string, IService> services = 
                HttpRuntime.Cache.Get("SERVICE_GALLERY") as IDictionary<string, IService>;
            IService service = null;

            if (services != null)
            {
                service = services[key];
            }

            return service;
        }
    }
}
