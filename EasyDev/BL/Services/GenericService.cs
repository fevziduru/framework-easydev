using System;
using System.Collections.Generic;
using EasyDev.PL;
using EasyDev.Resources;
using System.Data;
using WebCaching = System.Web.Caching;
using System.Web;
using EasyDev.SQMS;
using System.Threading;
using System.Web.Services;
using System.Linq.Expressions;
using System.Data.Common;
using EasyDev.BL.Services;
using EasyDev.Util;

namespace EasyDev.BL
{
        /// <summary>
        /// 通用服务,封装常用的数据操作
        /// </summary>
        public class GenericService : WebService, IService, IDataValidator
        {
                #region 私有成员

                /// <summary>
                /// 业务对象
                /// </summary>
                private IGenericBizObject bizObject = null;

                /// <summary>
                /// 业务对象工厂
                /// </summary>
                private BizObjectFactory boFactory = null;

                /// <summary>
                /// 本地服务管理器
                /// </summary>
                private NativeServiceManager serviceManager = null;

                /// <summary>
                /// 本地化/国际化对象创建工厂
                /// </summary>
                private ResourceManagerFactory resMgrFactory = null;

                /// <summary>
                /// 数据库会话缓存
                /// </summary>
                private WebCaching.Cache SessionPool
                {
                        get;
                        set;
                }

                /// <summary>
                /// 验证事件
                /// </summary>
                private event Func<DataSet, IDictionary<string, object>> Validation;

                #endregion

                /// <summary>
                /// 根据名称创建数据库SESSION
                /// </summary>
                /// <param name="name">配置文件中的SESSION名称</param>
                /// <returns></returns>
                public virtual GenericDBSession GetSession(string name)
                {
                        GenericDBSession session = SessionPool.Get(name + "_T" + Thread.CurrentThread.ManagedThreadId.ToString()) as GenericDBSession;


                        if (session == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(SessionPool, 100))
                                        {
                                                if (session == null)
                                                {
                                                        session = DBSessionManager.CreateDBSession(DataSourceManager.CreateDataSource(name)); //数据访问对象
                                                        session.SessionID = name + "_T" + Thread.CurrentThread.ManagedThreadId.ToString();
                                                        SessionPool.Insert(session.SessionID, session);
                                                }
                                        }
                                }
                                catch (ArgumentException arge)
                                {
                                        throw arge;
                                }
                                finally
                                {
                                        Monitor.Exit(SessionPool);
                                }
                        }

                        return session;
                }

                /// <summary>
                /// 取得公共库SESSION
                /// </summary>
                /// <returns></returns>
                public virtual GenericDBSession GetPubSession()
                {
                        return GetSession("SQMS.PUB");
                }

                /// <summary>
                /// 查找数据库SESSION,此方法根据表达式给出的条件到公共库中提取SESSION对象
                /// </summary>
                /// <param name="exp">查询表达式</param>
                /// <example>
                /// 从数据库中提取属于指定机构的数据库SESSION对象<br/>
                /// <code>
                /// GenericDBSession session = GetSession(p=>p["orgid"].ToString().Equal("spec_orgid"));
                /// </code>
                /// 取得与用户信息相关的数据库SESSION对象<br/>
                /// <code>
                /// GenericDBSession session = GetSession(p=>p.Field<string>("orgid").Equal(CurrentUser.OrganizationID));
                /// </code>
                /// </example>
                /// <returns></returns>
                public virtual GenericDBSession GetSession(Expression<Func<DataRow, bool>> exp)
                {
                        try
                        {
                                GenericDBSession pubsession = GetPubSession();
                                Func<DataRow, bool> func = exp.Compile();
                                DataSourceObject dso = null;
                                DataTable dtDso = pubsession.GetDataTableFromCommand(@"select db.connectionstring, db.provider from database db left join organization o on o.databaseid = db.databaseid");

                                foreach (DataRow row in dtDso.Rows)
                                {
                                        if ((bool)func.DynamicInvoke(row))
                                        {
                                                string provider = row.Field<string>("provider");
                                                string connstr = row.Field<string>("connectionstring");
                                                dso = new DataSourceObject(DbProviderFactories.GetFactory(provider));
                                                dso.ProviderName = row.Field<string>("name");
                                        }
                                }

                                return DBSessionManager.CreateDBSession(dso);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                

                #region 公共属性

                private UserInfo currentUser = null;

                /// <summary>
                /// 当前登录的用户信息
                /// </summary>
                public UserInfo CurrentUser
                {
                        get
                        {
                                UserIdentity ui = Thread.CurrentPrincipal.Identity as UserIdentity;
                                if (ui != null)
                                {
                                        if (ui.UserInfo != null)
                                        {
                                                return ui.UserInfo;
                                        }
                                        else
                                        {
                                                //if (ui.Name != null)
                                                //{
                                                //        PassportService srv = ServiceManager.CreateService<PassportService>();
                                                        
                                                //}
                                                //Session["USER_INFO"]

                                                return null;
                                        }
                                }
                                else
                                {
                                        return null;                                       
                                }
                        }
                        set
                        {
                                this.currentUser = value;
                        }
                }

                /// <summary>
                /// 本地化/国际化对象
                /// </summary>
                public ResourceManagerFactory ResourceManagerCreator
                {
                        get
                        {
                                return this.resMgrFactory;
                        }
                }

                /// <summary>
                /// 服务管理器
                /// </summary>
                public NativeServiceManager ServiceManager
                {
                        get
                        {
                                if (this.serviceManager == null)
                                {
                                        this.serviceManager = ServiceManagerFactory.CreateServiceManager<NativeServiceManager>();
                                }

                                return this.serviceManager;
                        }
                }

                /// <summary>
                /// 业务对象名称
                /// </summary>
                public string BOName
                {
                        get;
                        set;
                }

                /// <summary>
                /// 资源管理器
                /// </summary>
                public IResourceManager ResourceManager
                {
                        get;
                        set;
                }
                #endregion

                /// <summary>
                /// 会话名称
                /// </summary>
                public string SessionName
                {
                        get;
                        set;
                }

                /// <summary>
                /// 构造方法
                /// </summary>
                public GenericService()
                {
                        //准备验证事件
                        this.Validation += new Func<DataSet, IDictionary<string, object>>(Validate);

                        //初始化缓存
                        SessionPool = HttpRuntime.Cache;
                        
                        //服务初始化
                        Initialize();
                }

                /// <summary>
                /// 默认数据库会话对象，这个默认会话对象由EasyDev.Persistence.Config文件的DataSources结点的Default属性指定
                /// </summary>
                public GenericDBSession DefaultSession
                {
                        get
                        {
                                GenericDBSession session = SessionPool.Get("DEFAULT_SESSION_T" + Thread.CurrentThread.ManagedThreadId.ToString()) as GenericDBSession;


                                if (session == null)
                                {
                                        try
                                        {
                                                if (Monitor.TryEnter(SessionPool, 100))
                                                {
                                                        if (session == null)
                                                        {
                                                                session = DBSessionManager.CreateDBSession(DataSourceManager.CreateDefaultDataSource());     //数据访问对象
                                                                session.SessionID = "DEFAULT_SESSION_T" + Thread.CurrentThread.ManagedThreadId.ToString();
                                                                SessionPool.Insert(session.SessionID, session);
                                                        }
                                                }
                                        }
                                        catch (ArgumentException arge)
                                        {
                                                throw arge;
                                        }
                                        finally
                                        {
                                                Monitor.Exit(SessionPool);
                                        }
                                }

                                return session;
                        }
                }

                /// <summary>
                /// 服务初始化
                /// TODO:SESSION名称待重构
                /// </summary>
                protected virtual void Initialize()
                {
                        //SESSION名称由可视名称和线程ID组成
                        string sessionName = (SessionName == null || SessionName.Equals(string.Empty)) ? "DEFAULT_SESSION" : SessionName;

                        this.resMgrFactory = ResourceManagerFactory.GetInstance();      //国际化资源对象
                        this.serviceManager = ServiceManagerFactory.CreateServiceManager<NativeServiceManager>();

                        //创建默认SESSION
                        GenericDBSession session = SessionPool.Get(sessionName + "_T" + Thread.CurrentThread.ManagedThreadId.ToString()) as GenericDBSession;


                        if (session == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(SessionPool, 100))
                                        {
                                                if (session == null)
                                                {
                                                        session = DBSessionManager.CreateDBSession(DataSourceManager.CreateDefaultDataSource()); //数据访问对象
                                                        session.SessionID = sessionName + "_T" + Thread.CurrentThread.ManagedThreadId.ToString();
                                                        SessionPool.Insert(session.SessionID, session);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(SessionPool);
                                }
                        }

                        this.boFactory = BizObjectFactory.GetInstance(session);   //BO工厂
                        if (this.BOName == null || this.BOName.Length == 0)
                        {
                                throw new ArgumentException("BOName属性不能为空", "BOName");
                        }

                        this.bizObject = boFactory.CreateBizObject(this.BOName);   //创建BO对象
                }

                /// <summary>
                /// 根据单主键加载数据
                /// </summary>
                /// <param name="key"></param>
                /// <returns></returns>
                public virtual DataSet LoadByKey(object key)
                {
                        return this.LoadByKey(key, false);
                }

                /// <summary>
                /// 根据单主键加载数据
                /// </summary>
                /// <param name="key">主键值</param>
                /// <param name="isCreateNew">是否创建新行</param>
                /// <returns></returns>
                public virtual DataSet LoadByKey(object key, bool isCreateNew)
                {
                        try
                        {
                                this.bizObject.LoadByPrimaryKey(key);
                                if (this.bizObject.IsEmpty())
                                {
                                        if (isCreateNew)
                                        {
                                                this.bizObject.Create();
                                        }
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }

                        return this.bizObject.BOData;
                }

                /// <summary>
                /// 根据多主键加载数据
                /// </summary>
                /// <param name="key"></param>
                /// <returns></returns>
                public virtual DataSet LoadByKeys(IDictionary<string, object> keys)
                {
                        return this.LoadByKeys(keys, false);
                }

                /// <summary>
                /// 根据多主键加载数据
                /// </summary>
                /// <param name="keys">主键值集合</param>
                /// <param name="isCreateNew">是否创建新行</param>
                /// <returns></returns>
                public virtual DataSet LoadByKeys(IDictionary<string, object> keys, bool isCreateNew)
                {
                        try
                        {
                                if (keys != null)
                                {
                                        this.bizObject.LoadByPrimaryKeys(keys);
                                }
                                if (this.bizObject.IsEmpty())
                                {
                                        if (isCreateNew)
                                        {
                                                this.bizObject.Create();
                                        }
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }

                        return this.bizObject.BOData;
                }

                /// <summary>
                /// 创建新行
                /// </summary>
                /// <returns></returns>
                public virtual DataRow Create()
                {
                        return this.bizObject.Create();
                }

                /// <summary>
                /// 通用保存方法，保存与当前BOName指定的数据表相关的数据
                /// </summary>
                /// <param name="dsSave"></param>
                public virtual void Save(DataSet dsSave)
                {
                        #region 数据验证事件
                        IDictionary<string, object> result = null;
                        try
                        {
                                result = Validation(dsSave);
                                if (((bool)result[ValidatorResultKey.SUCCESS]) == false)
                                {
                                        throw new ServiceValidationException(result[ValidatorResultKey.ERRMSG].ParseString());
                                }
                        }
                        catch (ServiceValidationException e)
                        {
                                throw e;
                        } 
                        #endregion

                        this.bizObject.BOData = dsSave;
                        this.bizObject.Synchronize();
                }

                /// <summary>
                /// 加载所有数据
                /// </summary>
                /// <returns></returns>
                public virtual DataSet LoadAll()
                {
                        try
                        {
                                if (this.bizObject.LoadAll())
                                {
                                        return this.bizObject.BOData;
                                }
                                else
                                {
                                        throw new Exception();
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 加载视图
                /// </summary>
                /// <param name="name"></param>
                /// <returns></returns>
                public virtual DataSet LoadViewByName(string name)
                {
                        try
                        {
                                return this.bizObject.LoadViewByName(name);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 根据主键值删除数据
                /// </summary>
                /// <param name="key"></param>
                public virtual void DeleteByKey(object key)
                {
                        try
                        {
                                this.bizObject.DeleteByPrimaryKey(key);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 删除由多主属性构成的数据表数据
                /// </summary>
                /// <param name="keys"></param>
                public virtual void DeleteByKeys(IDictionary<string, object> keys)
                {
                        try
                        {
                                this.bizObject.DeleteByPrimaryKeys(keys);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 根据条件删除数据
                /// </summary>
                /// <param name="condition">删除条件，语法同SQL中的WHERE，但不带WHERE关键字</param>
                public virtual void DeleteByCondition(string condition)
                {
                        try
                        {
                                this.bizObject.DeleteByCondition(condition);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 按条件加载
                /// </summary>
                /// <param name="cond">加载条件，语法同SQL中的WHERE，但不带WHERE关键字</param>
                /// <returns></returns>
                public virtual DataSet LoadByCondition(string cond)
                {
                        try
                        {
                                if (this.bizObject.LoadByCondition(cond))
                                {
                                        return this.bizObject.BOData;
                                }
                                else
                                {
                                        throw new Exception();
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 取下一个ORACLE序列号
                /// </summary>
                /// <returns></returns>
                public string GetNextSequenceID()
                {
                        return this.bizObject.GetNextSequenceID(this.BOName);
                }

                /// <summary>
                /// 取得引用表字段数据
                /// </summary>
                /// <param name="session">数据库会话对象</param>
                /// <param name="refTextField">引用的文本字段</param>
                /// <param name="refTable">引用数据表</param>
                /// <param name="refKey">引用字段</param>
                /// <param name="refValue">引用数据值</param>
                /// <returns></returns>
                public virtual object GetReferenceValue(
                    GenericDBSession session,
                    string refTextField,
                    string refTable,
                    string refKey,
                    string refValue)
                {
                        try
                        {
                                return session.GetScalarObjectFromCommand(
                                    string.Format(@"select {0} from {1} where {2}=:value", refTextField, refTable, refKey), refValue);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 取得引用表字段数据
                /// </summary>
                /// <param name="refTextField">引用的文本字段</param>
                /// <param name="refTable">引用数据表</param>
                /// <param name="refKey">引用字段</param>
                /// <param name="refValue">引用数据值</param>
                /// <returns></returns>
                public virtual object GetReferenceValue(string refTextField,
                    string refTable,
                    string refKey,
                    string refValue)
                {
                        try
                        {
                                return this.DefaultSession.GetScalarObjectFromCommand(
                                    string.Format(@"select {0} from {1} where {2}=:value", refTextField, refTable, refKey), refValue);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 生成编号
                /// </summary>
                /// <returns></returns>
                public virtual string GenerateCode()
                {
                        return this.BOName.Substring(0, 2) + "-" + DateTime.Now.Ticks.ToString();
                }

                #region IDataValidator 成员

                /// <summary>
                /// 默认验证方法
                /// </summary>
                /// <param name="param"></param>
                /// <returns></returns>
                public virtual IDictionary<string, object> Validate(DataSet param)
                {
                        IDictionary<string, object> result = new Dictionary<string, object>();
                        result.Add(ValidatorResultKey.SUCCESS, true);

                        DataTable dtBO = param.Table(BOName);
                        foreach (DataRow row in dtBO.Rows)
                        {
                                foreach (DataColumn col in dtBO.Columns)
                                {
                                        //Date列的长度是-1
                                        if (col.MaxLength < row[col.ColumnName].ParseString().Length && col.MaxLength > -1)
                                        {
                                                result[ValidatorResultKey.SUCCESS] = false;
                                                result[ValidatorResultKey.ERRMSG] = col.ColumnName + "列的数据长度太大";
                                        }
                                }
                        }

                        return result;
                }

                #endregion
        }

        public static class ValidatorResultKey
        {
                public static readonly string SUCCESS = "SUCCESS";
                public static readonly string ERRMSG = "ERRMSG";
        }
}
