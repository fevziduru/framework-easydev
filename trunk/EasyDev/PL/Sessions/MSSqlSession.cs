using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using EasyDev.Util;
using System;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据库连接会话，会话对象通过获取SQL命令文本执行SQL命令，
    /// 也可以从外部获取继承于IDbCommand接口的对象来执行SQL命令
    /// </summary>
    [Obsolete("此类已经过时")]
    public class MSSqlSession : DBSessionBase, IDBSession
    {
        /// <summary>
        /// 默认构造器
        /// </summary>
        public MSSqlSession()
        {
            this._adapter = new SqlDataAdapter();

            this.dataSourceName = "MSSqlDataSource";

            this.InitDataSource();
        }

        /// <summary>
        /// 构造器
        /// 根据数据源名称初始化数据库会话
        /// </summary>
        /// <param name="dataSourceName">数据源名称</param>
        public MSSqlSession(string dataSourceName)
        {
            this._adapter = new SqlDataAdapter();

            this.dataSourceName = dataSourceName;

            this.InitDataSourceByName(dataSourceName);
        }

        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public override string ConnectionString
        {
            get
            {
                //如果连接字符串为空则从配置文件中取
                if (connectionString.Equals(""))
                { 
                    connectionString = ConvertUtil.ConvertToString(ConfigurationManager.ConnectionStrings[this.dataSourceName]);
                }

                //如果没有指定数据连接串配置名称，则默认取第一个
                if (connectionString.Equals(""))
                {
                    connectionString = ConvertUtil.ConvertToString(ConfigurationManager.ConnectionStrings[0]);
                }

                return connectionString;
            }
            set
            {
                this.ConnectionString = value;
            }
        }

        #region IDBSession 成员

        /// <summary>
        /// 初始化数据源
        /// </summary>
        public override void InitDataSource()
        {
            try
            {
                this._connection = new SqlConnection(this.ConnectionString);
                this._command = new SqlCommand();
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 根据配置名称初始化数据源
        /// </summary>
        /// <param name="cfgName">数据源配置节名称</param>
        public override void InitDataSourceByName(string cfgName)
        {
            try
            {
                string customConnectionString = ConvertUtil.ConvertToString(ConfigurationManager.ConnectionStrings[cfgName]);
                this._connection = new SqlConnection(customConnectionString);
                this._command = new SqlCommand();
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 根据SQL创建命令对象
        /// </summary>
        /// <param name="sqlCmdText">SQL命令文本</param>
        public override IDbCommand CreateCommand(string sqlCmdText) 
        {
            IDbCommand command = null;
            try
            {
                if (null == command)
                {
                    command = new SqlCommand(sqlCmdText,(SqlConnection)this._connection);
                }
                else
                {
                    command.CommandText = sqlCmdText;
                    command.Connection = this._connection;
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }

            return command;
        }
        #endregion

        /// <summary>
        /// 开启事务
        /// </summary>
        public override void BeginTransaction()
        {
            try
            {
                if (this._connection != null)
                {
                    this._transaction = this._connection.BeginTransaction(IsolationLevel.Serializable);
                    if (this._command == null)
                    {
                        this._command = new SqlCommand();
                    }

                    this._command.Transaction = this._transaction;             //初始化命令对象的事务属性
                }
                else
                {
                    throw new DALCoreException("Failed_to_begin_transaction");      //开启事务失败,可能是数据库连接没有被初始化
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }
        
        #region 待实现方法

        /// <summary>
        /// 获取存储过程命令封装器
        /// </summary>
        /// <returns></returns>
        //public override IDBCommandWrapper GetStoredProcCommandWrapper()
        //{
        //    throw new Exception("The method or operation is not implemented.");
        //}

        /// <summary>
        /// 获取参数化查询对象
        /// </summary>
        /// <param name="sqlCmd">参数化查询字串</param>
        /// <param name="parameters">参数列表(键值对)</param>
        /// <returns></returns>
        //public override IDBCommandWrapper GetSqlCommandWithParametersWrapper(string sqlCmd, IDictionary<string, object> parameters)
        //{
        //    try
        //    {
        //        if (this._command != null)
        //        {
        //            this._command.CommandText = sqlCmd;
        //        }
        //        else
        //        {
        //            throw new DALCoreException("_command_object_is_null");     //命令对象不存在
        //        }

        //        return new DBCommandWrapperBase(this._command, parameters);
        //    }
        //    catch (DALCoreException e)
        //    {
        //        throw e;
        //    }
        //}

        #endregion
    }
}
