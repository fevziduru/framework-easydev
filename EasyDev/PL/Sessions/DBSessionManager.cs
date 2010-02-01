using EasyDev.Configuration;
using EasyDev.PL;
using System.Xml;
using System.Collections;
using System;
using System.Collections.Generic;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据库会话管理器，用于通过配置创建数据库会话对象
    /// </summary>
    public class DBSessionManager
    {   
        #region IDBSessionManager 成员

        /// <summary>
        /// 根据用于初始化SessionManager对象的Provider产生数据库会话对象
        /// </summary>
        /// <param name="_datasource">datasource名称</param>
        /// <returns></returns>
        public static GenericDBSession CreateDBSession(DataSourceObject _datasource)
        {
            try
            {
                return new GenericDBSession(_datasource, PersistenceConfigManager.GetInstance());
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        #endregion
    }
}
