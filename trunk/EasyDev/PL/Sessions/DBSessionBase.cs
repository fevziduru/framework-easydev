using System.Data;
using EasyDev.PL;
using System;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据库连接的公共基类
    /// </summary>
    [Obsolete("此类已经过时")]
    public abstract class DBSessionBase  : IDBSession
    {
        /// <summary>
        /// 数据源名称
        /// </summary>
        protected string dataSourceName = "";

        /// <summary>
        /// 连接字符串
        /// </summary>
        protected string connectionString = "";

        /// <summary>
        /// 数据库连接对象
        /// </summary>
        protected IDbConnection _connection = null;

        /// <summary>
        /// 数据库命令对象
        /// </summary>
        protected IDbCommand _command = null;

        /// <summary>
        /// 数据库事务对象
        /// </summary>
        protected IDbTransaction _transaction = null;

        /// <summary>
        /// 数据库适配器
        /// </summary>
        protected IDbDataAdapter _adapter = null;
        
        /// <summary>
        /// 连接字符串
        /// </summary>
        public virtual string ConnectionString
        {
            get
            {
                return this.connectionString;
            }
            set
            {
                this.connectionString = value;
            }
        }

        /// <summary>
        /// 数据适配器
        /// </summary>
        public IDbDataAdapter DataAdapter
        {
            get
            {
                return this._adapter;
            }
        }

        /// <summary>
        /// 返回数据库连接
        /// </summary>
        public IDbConnection Connection
        {
            get 
            {
                if (this._connection == null)
                {
                    throw new DALCoreException("the_connection_is_not_initialized");
                }
                return this._connection;
            }
        }

        /// <summary>
        /// 取得数据库会话状态
        /// </summary>
        public virtual ConnectionState SessionState
        {
            get
            {
                if (this._connection != null)
                {
                    return this._connection.State;
                }
                else
                {
                    throw new DALCoreException("");
                }
            }
        }

        /// <summary>
        /// 开启数据库连接会话
        /// </summary>
        public virtual void OpenSession()
        {
            try
            {
                if (this._connection.State == ConnectionState.Closed)
                {
                    this._connection.Open();
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 关闭数据库会话
        /// </summary>
        public virtual void CloseSession()
        {
            try
            {
                if (this._connection.State == ConnectionState.Open)
                {
                    this._connection.Close();
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }  

        /// <summary>
        /// 提交事务
        /// </summary>
        public virtual void CommitTransaction()
        {
            try
            {
                if (this._transaction != null)
                {
                    this._transaction.Commit();
                }
                else
                {
                    throw new DALCoreException("command_is_null_transaction");    //事务提交失败,可能是命令对象未初始化
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 回滚事务
        /// </summary>
        public virtual void RollBackTransaction()
        {
            try
            {
                if (this._transaction != null)
                {
                    this._transaction.Rollback();
                }
                else
                {
                    throw new DALCoreException("command_is_null");     //命令对象为空
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 执行SQL命令
        /// </summary>
        /// <returns>SQL命令影响行数</returns>
        public virtual int ExecuteCommand(string sqlCmd)
        {
            int result = -1;

            try
            {
                if (this._command != null)
                {
                    this._command.CommandText = sqlCmd;
                    if (this._connection != null)
                    {
                        this._command.Connection = this._connection;
                    }
                    else
                    {
                        throw new DALCoreException("command_no_connection");    //命令对象没有对应的数据库连接信息
                    }

                    this._connection.Open();                                                            //开启数据库连接
                    this.BeginTransaction();                                                             //开启事务
                    
                    result = this._command.ExecuteNonQuery();

                    this._transaction.Commit();
                }
                else
                {
                    throw new DALCoreException("command_is_null");     //命令对象为空
                }
            }
            catch (DALCoreException e)
            {
                this._transaction.Rollback();                                               //回滚事务
                throw e;
            }
            finally
            {
                this._connection.Close();                                                    //关闭连接
            }

            return result;
        }

        /// <summary>
        /// 执行SQL命令
        /// </summary>
        /// <returns>SQL命令影响行数</returns>
        public virtual int ExecuteCommand(IDbCommand command)
        {
            int result = -1;
            try
            {
                if (string.IsNullOrEmpty(command.CommandText) == false && command.Connection != null)
                {
                    this._connection.Open();                                                     //开启数据库连接
                    this._transaction = this._connection.BeginTransaction();       //开启事务
                    command.Transaction = this._transaction;                           //初始化Command对象的事务属性

                    result = command.ExecuteNonQuery();

                    this._transaction.Commit();
                }
                else
                {
                    throw new DALCoreException("no_connection_or_no_commandText"); //命令对象没有命令文本或没有初始化数据库连接
                }
            }
            catch (DALCoreException e)
            {
                this._transaction.Rollback();   //回滚事务
                throw e;
            }
            finally
            {
                this._connection.Close();      //关闭数据库连接
            }
            return result;
        }
               
        /// <summary>
        /// 根据SQL命令得到一个数据记录游标
        /// </summary>
        /// <returns></returns>
        public virtual IDataReader GetDataReaderFromCommand(string sqlCmd)
        {
            try
            {
                if (this._command != null)
                {
                    this._command.CommandText = sqlCmd;
                    this._command.Connection = this.Connection;
                    
                    this.OpenSession();

                    return this._command.ExecuteReader(CommandBehavior.CloseConnection);
                }
                else
                {
                    throw new DALCoreException("command_is_null");     //命令对象为空
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
            finally
            {
                this.CloseSession();
            }
        }

        /// <summary>
        /// 根据命令对象得到一个数据记录游标
        /// </summary>
        /// <returns></returns>
        public virtual IDataReader GetDataReaderFromCommand(IDbCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.CommandText) == false && command.Connection != null)
                {
                    this.OpenSession();
                    return command.ExecuteReader(CommandBehavior.CloseConnection);
                }
                else
                {
                    throw new DALCoreException("no_connection_or_no_commandText"); //命令对象没有命令文本或没有初始化数据库连接
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
            finally
            {
                this.CloseSession();
            }
        }

        /// <summary>
        /// 取得结果集中第一行第一列的值，以对象类型返回
        /// </summary>
        /// <param name="sqlCmd">SQL命令文本</param>
        /// <returns></returns>
        public virtual object GetScalarObjectFromCommand(string sqlCmd)
        {
            try
            {
                if (this._command != null)
                {
                    this._command.CommandText = sqlCmd;
                    this._command.Connection = this.Connection;
                    
                    this.OpenSession();

                    return this._command.ExecuteScalar();
                }
                else
                {
                    throw new DALCoreException("command_is_null");     //命令对象为空
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
            finally
            {
                this.CloseSession();
            }
        }

        /// <summary>
        /// 取得结果集中第一行第一列的值，以对象类型返回
        /// </summary>
        /// <param name="sqlCmd">SQL命令文本</param>
        /// <returns></returns>
        public virtual object GetScalarObjectFromCommand(IDbCommand command)
        {
            try
            {
                if (string.IsNullOrEmpty(command.CommandText) == false && command.Connection != null)
                {
                    this.OpenSession();

                    return command.ExecuteScalar();
                }
                else
                {
                    throw new DALCoreException("no_connection_or_no_commandText"); //命令对象没有命令文本或没有初始化数据库连接
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
            finally
            {
                this.CloseSession();
            }
        }

        /// <summary>
        /// 获取数据集
        /// </summary>
        /// <param name="sqlCmd">SQL命令文本</param>
        /// <returns></returns>
        public virtual DataSet GetDataSetFromCommand(string sqlCmd)
        {
            DataSet dsResult = new DataSet();
            try
            {
                if (this._command != null)
                {
                    this._command.CommandText = sqlCmd;
                    this._command.Connection = this.Connection;

                    if (this._adapter != null)
                    {
                        this._adapter.SelectCommand = this._command;
                        this._adapter.Fill(dsResult);
                    }
                    else
                    {
                        throw new DALCoreException("the_adapter_is_not_initialized");        //Adapter未初始化
                    }
                }
                else
                {
                    throw new DALCoreException("command_is_null");     //命令对象为空
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }

            return dsResult;
        }

        /// <summary>
        /// 获取数据集
        /// </summary>
        /// <param name="command">命令对象</param>
        /// <returns></returns>
        public virtual DataSet GetDataSetFromCommand(IDbCommand command)
        {
            DataSet dsResult = new DataSet();

            try
            {
                if (string.IsNullOrEmpty(command.CommandText) == false && command.Connection != null)
                {
                    if (this._adapter != null)
                    {
                        this._adapter.SelectCommand = command;
                        this._adapter.Fill(dsResult);
                    }
                    else
                    {
                        throw new DALCoreException("the_adapter_is_not_initialized");        //Adapter未初始化
                    }
                }
                else
                {
                    throw new DALCoreException("no_connection_or_no_commandText"); //命令对象没有命令文本或没有初始化数据库连接
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }

            return dsResult;
        }

        /// <summary>
        /// 获取数据表
        /// </summary>
        /// <param name="command">命令对象</param>
        /// <returns></returns>
        public virtual DataTable GetDataTableFromCommand(IDbCommand command)
        {
            DataTable dtResult = new DataTable();

            try
            {
                DataSet dsTmp = this.GetDataSetFromCommand(command);
                dtResult = dsTmp.Tables[0];
                dsTmp.Tables.Clear();
            }
            catch (DALCoreException e)
            {
                throw e;
            }

            return dtResult;
        }

        /// <summary>
        /// 获取数据表
        /// </summary>
        /// <param name="sqlCmd">命令文本</param>
        /// <returns></returns>
        public virtual DataTable GetDataTableFromCommand(string sqlCmd)
        {
            DataTable dtResult = new DataTable();

            try
            {
                if (this._command != null)
                {
                    this._command.CommandText = sqlCmd;
                    this._command.Connection = this.Connection;
                    dtResult = this.GetDataTableFromCommand(this._command);
                }
                else
                {
                    throw new DALCoreException("the_command_is_not_initialized");  //为初始化命令对象
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }

            return dtResult;
        }

        /// <summary>
        /// 取得存储过程命令封装器
        /// </summary>
        /// <returns></returns>
        //public abstract IDBCommandWrapper<> GetStoredProcCommandWrapper();

        /// <summary>
        /// 取得SQL命令文本封装器
        /// </summary>
        /// <returns></returns>
        //public abstract IDBCommandWrapper GetSqlCommandWithParametersWrapper(string sqlCmd ,IDictionary<string,object> parameters);

        /// <summary>
        /// 开启数据库事务
        /// </summary>
        public abstract void BeginTransaction();

        #region IDBSession 成员

        public virtual void InitDataSource()
        {
            throw new System.Exception("The method or operation is not implemented.");
        }

        public virtual void InitDataSourceByName(string cfgName)
        {
            throw new System.Exception("The method or operation is not implemented.");
        }

        public virtual IDbCommand CreateCommand(string sqlCmdText)
        {
            throw new System.Exception("The method or operation is not implemented.");
        }

        #endregion
    }
}