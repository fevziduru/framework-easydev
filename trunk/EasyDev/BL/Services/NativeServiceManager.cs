using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.SQMS;

namespace EasyDev.BL
{
    /// <summary>
    /// 本地服务管理器
    /// </summary>
    public class NativeServiceManager : AbstractServiceManager
    {
        public NativeServiceManager()
        { }

        /// <summary>
        /// 创建本地服务
        /// </summary>
        /// <typeparam name="T">服务类型</typeparam>
        /// <returns></returns>
        public T CreateService<T>() where T : IService, new()
        {
            return new T();
        }

        /// <summary>
        /// 创建服务，并创建sessionName指定的数据库Session对象
        /// </summary>
        /// <typeparam name="T">服务类型</typeparam>
        /// <param name="sessionName">数据库会话对象</param>
        /// <returns>指定类型的服务</returns>
        public T CreateService<T>(string sessionName) where T : IService, new()
        {
            T srv = new T();
            srv.SessionName = sessionName;

            return srv;
        }

        /// <summary>
        /// 创建服务，并将指定服务的用户信息复制到新的服务
        /// </summary>
        /// <typeparam name="T">服务类型</typeparam>
        /// <param name="outsrv">源服务</param>
        /// <returns>指定类型的服务</returns>
        public T CreateService<T>(T outsrv) where T : GenericService, new()
        {
            T srv = new T();
            srv.CurrentUser = outsrv.CurrentUser.Clone() as UserInfo;

            return srv;
        }

        /// <summary>
        /// 实例化服务对象
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public IService CreateService(Type type)
        {
            return Activator.CreateInstance(type) as IService;
        }
    }
}