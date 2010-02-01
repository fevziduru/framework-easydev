using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.BL
{
    /// <summary>
    /// 服务管理器工厂
    /// </summary>
    public class ServiceManagerFactory
    {
        /// <summary>
        /// 创建管理器
        /// </summary>
        /// <returns></returns>
        public static T CreateServiceManager<T>() where T : IServiceManager, new()
        {
            return new T();
        }

        /// <summary>
        /// 创建远程服务管理器
        /// </summary>
        /// <returns></returns>
        public static RemotingServiceManager CreateRemotingServiceManager()
        {
            return new RemotingServiceManager();
        }

        /// <summary>
        /// 创建本地服务管理器
        /// </summary>
        /// <returns></returns>
        public static NativeServiceManager CreateNativeServiceManager()
        {
            return new NativeServiceManager();
        }
    }
}
