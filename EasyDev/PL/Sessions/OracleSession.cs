using System;
using System.Data;
using System.Data.OracleClient;
using System.Configuration;
using EasyDev.Util;

namespace EasyDev.PL
{
    [Obsolete("此类已经过时")]
    public class OracleSession : DBSessionBase, IDBSession
    {
        /// <summary>
        /// 默认构造器
        /// </summary>
        public OracleSession()
        {
            this._adapter = new OracleDataAdapter();

            this.dataSourceName = "OracleDataSource";

            this.InitDataSource();
        }

        /// <summary>
        /// 构造器
        /// 根据数据源名称初始化数据库会话
        /// </summary>
        /// <param name="_dataSourceName">数据源名称</param>
        public OracleSession(string _dataSourceName)
        {
            this.dataSourceName = _dataSourceName;

            this.InitDataSourceByName(_dataSourceName);
        }

        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        public override string ConnectionString
        {
            get
            {
                //如果连接字符串为空则从配置文件中取
                if (this.connectionString.Equals(""))
                {
                    this.connectionString = ConvertUtil.ConvertToString(ConfigurationManager.ConnectionStrings[this.dataSourceName]);
                }

                //如果没有指定数据连接串配置名称，则默认取第一个
                if (this.connectionString.Equals(""))
                {
                    this.connectionString = ConvertUtil.ConvertToString(ConfigurationManager.ConnectionStrings[0]);
                }

                return this.connectionString;
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
                this._connection = new OracleConnection(this.ConnectionString);
                this._command = new OracleCommand();
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
                string customConnectionString = System.Convert.ToString(ConfigurationManager.ConnectionStrings[cfgName]);
                this._connection = new OracleConnection(customConnectionString);
                this._command = new OracleCommand();
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
                    command = new OracleCommand(sqlCmdText, (OracleConnection)this._connection);
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
                        this._command = new OracleCommand();
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

        //public override IDBCommandWrapper<> GetStoredProcCommandWrapper()
        //{
        //    throw new Exception("The method or operation is not implemented.");
        //}

        //public override IDBCommandWrapper GetSqlCommandWithParametersWrapper(string sqlCmd, IDictionary<string, object> parameters)
        //{
        //    throw new Exception("The method or operation is not implemented.");
        //}

        #endregion
    }
}