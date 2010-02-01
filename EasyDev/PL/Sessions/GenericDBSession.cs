using System.Data;
using System.Data.Common;
using System;
using System.Text.RegularExpressions;
using EasyDev.Configuration;
using System.Collections.Generic;
using EasyDev.Util;

namespace EasyDev.PL
{
        /// <summary>
        /// 数据库连接的公共基类
        /// </summary>
        public class GenericDBSession : IGenericDBSession
        {
                private string sessionId = "";

                #region IGenericDBSession 成员

                public string SessionID
                {
                        get
                        {
                                return sessionId;
                        }
                        set
                        {
                                sessionId = value;
                        }
                }

                #endregion
                /// <summary>
                /// 数据源名称
                /// </summary>
                protected string dataSourceName = "";

                /// <summary>
                /// 连接字符串 
                /// </summary>
                protected string connectionString = "";

                /// <summary>
                /// 数据库事务对象
                /// </summary>
                protected DbTransaction _transaction = null;

                /// <summary>
                /// 主键值生成器
                /// </summary>
                public IGenerator IdentityGenerator
                {
                        get;
                        set;
                }

                private PersistenceConfigManager ConfigManager
                {
                        get;
                        set;
                }

                /// <summary>
                /// 数据源对象
                /// </summary>
                private DataSourceObject DataSource
                {
                        get;
                        set;
                }

                /// <summary>
                /// 当前使用的数据库提供程序
                /// </summary>
                public string CurrentDbProvider
                {
                        get;
                        set;
                }

                /// <summary>
                /// 连接字符串
                /// </summary>
                public virtual string ConnectionString
                {
                        get;
                        set;
                }

                /// <summary>
                /// 数据适配器
                /// </summary>
                public DbDataAdapter DataAdapter
                {
                        get;
                        set;
                }

                private DbConnection _connection;

                /// <summary>
                /// 返回数据库连接
                /// </summary>
                public DbConnection Connection
                {
                        get
                        {
                                if (this._connection == null)
                                {
                                        throw new DALCoreException("the_connection_is_not_initialized");
                                }
                                return this._connection;
                        }
                        set
                        {
                                this._connection = value;
                        }
                }

                /// <summary>
                /// 通用数据命令对象
                /// </summary>
                public DbCommand Command
                {
                        get;
                        set;
                }

                /// <summary>
                /// 取得数据库会话状态
                /// </summary>
                public virtual ConnectionState SessionState
                {
                        get
                        {
                                if (this.Connection != null)
                                {
                                        return this.Connection.State;
                                }
                                else
                                {
                                        throw new DALCoreException("");
                                }
                        }
                }

                /// <summary>
                /// 构造方法
                /// </summary>
                /// <param name="_datasource"></param>
                public GenericDBSession(DataSourceObject _datasource, PersistenceConfigManager _cfg)
                {
                        DataSource = _datasource;
                        ConfigManager = _cfg;

                        Connection = DataSource.GetDBConnection();
                        DataAdapter = DataSource.GetDBDataAdapter();
                        Command = DataSource.GetDBCommand();

                        foreach (IDataSource providerItem in ConfigManager.DataSources.Values)
                        {
                                if (DataSource.ProviderName == providerItem.Name)
                                {
                                        this.Connection.ConnectionString = providerItem.ConnectionString;
                                        CurrentDbProvider = providerItem.ProviderType;
                                        break;
                                }
                        }

                        if (CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
                                CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                        {
                                IdentityGenerator = GeneratorFactory.CreateGenerator<OracleSequenceGenerator>();
                                ((OracleSequenceGenerator)IdentityGenerator).TableName = "";
                                ((OracleSequenceGenerator)IdentityGenerator).Session = this;
                        }
                        else
                        {
                                IdentityGenerator = GeneratorFactory.CreateGenerator<CommonIdentityGenerator>();
                        }
                }

                /// <summary>
                /// 开启数据库连接会话
                /// </summary>
                public virtual void OpenSession()
                {
                        OpenSession(false);
                }

                /// <summary>
                /// 开启数据库会话
                /// </summary>
                /// <param name="withTrans">为TRUE表示在开启会话的时候同时开始事务</param>
                public virtual void OpenSession(bool withTrans)
                {
                        try
                        {
                                if (this.Connection.State == ConnectionState.Closed)
                                {
                                        this.Connection.Open();
                                        if (withTrans)
                                        {
                                                this._transaction = this.Connection.BeginTransaction();
                                        }
                                }
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 设置参数化查询命令
                /// </summary>
                /// <param name="commandText">参数化命令</param>
                /// <param name="paramValues">参数值列表</param>
                public virtual void PrepareParameterizedCommand(string commandText, params object[] paramValues)
                {
                        try
                        {
                                MatchCollection collection = null;

                                if (CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
                                    CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                                {
                                        collection = Regex.Matches(commandText, "[^\\w*]:\\w+", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                                }
                                else
                                {
                                        collection = Regex.Matches(commandText, "@\\w+", RegexOptions.Multiline | RegexOptions.IgnoreCase);
                                }

                                if (collection.Count == 0)
                                {
                                        throw new ArgumentException("commandText_is_not_a_parameterized_query");
                                }

                                if (collection.Count > paramValues.Length)
                                {
                                        throw new ArgumentException("too_many_parameters");
                                }

                                Command.CommandText = commandText;

                                for (int i = 0; i < collection.Count; i++)
                                {

                                        DbParameter parameter = DataSource.GetDBParameter();
                                        parameter.Direction = ParameterDirection.Input;
                                        parameter.ParameterName = collection[i].ToString().Substring(collection[i].ToString().IndexOf(":"));
                                        parameter.Value = paramValues[i];

                                        if (CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
                                            CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                                        {
                                                parameter.SourceColumn = collection[i].ToString().Substring(collection[i].ToString().IndexOf(":") + 1);
                                        }
                                        else
                                        {
                                                parameter.SourceColumn = collection[i].ToString().Replace("@", "");
                                        }

                                        Command.Parameters.Add(parameter);
                                }
                        }
                        catch (ArgumentException e)
                        {
                                ResetCommand();
                                throw e;
                        }
                }

                /// <summary>
                /// 设置参数化查询命令
                /// </summary>
                /// <param name="commandText"></param>
                /// <param name="paramValues"></param>
                public virtual void PrepareParameterizedCommand(string commandText, IDictionary<string, object> paramValues)
                {
                        try
                        {
                                MatchCollection collection = null;

                                if (CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
                                    CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                                {
                                        collection = Regex.Matches(commandText, "[^\\w*]:\\w+", RegexOptions.Multiline);
                                }
                                else
                                {
                                        collection = Regex.Matches(commandText, "@\\w+", RegexOptions.Multiline);
                                }

                                if (collection.Count == 0)
                                {
                                        throw new ArgumentException("commandText_is_not_a_parameterized_query");
                                }

                                if (collection.Count > paramValues.Count)
                                {
                                        throw new ArgumentException("too_many_parameters");
                                }

                                Command.CommandText = commandText;

                                for (int i = 0; i < collection.Count; i++)
                                {

                                        DbParameter parameter = DataSource.GetDBParameter();
                                        parameter.Direction = ParameterDirection.Input;
                                        parameter.ParameterName = collection[i].ToString();

                                        if (CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
                                            CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                                        {
                                                parameter.SourceColumn = collection[i].ToString().Replace(":", "");
                                        }
                                        else
                                        {
                                                parameter.SourceColumn = collection[i].ToString().Replace("@", "");
                                        }
                                        parameter.Value = paramValues[parameter.SourceColumn];

                                        Command.Parameters.Add(parameter);
                                }
                        }
                        catch (ArgumentException e)
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
                                if (this.Connection.State == ConnectionState.Open)
                                {
                                        this.Connection.Close();
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
                                if (this.Command != null)
                                {
                                        this.Command.CommandText = sqlCmd;
                                        if (this.Connection != null)
                                        {
                                                this.Command.Connection = this.Connection;
                                        }
                                        else
                                        {
                                                throw new DALCoreException("command_no_connection");    //命令对象没有对应的数据库连接信息
                                        }

                                        this.OpenSession(true);
                                        this.Command.Transaction = this._transaction;    //初始化命令对象的事务属性
                                        result = this.Command.ExecuteNonQuery();

                                        if (this._transaction != null)
                                        {
                                                this._transaction.Commit();
                                        }
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
                                ResetCommand();
                                this.Connection.Close();                                                    //关闭连接
                        }

                        return result;
                }

                /// <summary>
                /// 执行参数化查询，如果命令中不带参数则抛出ArgumentException异常
                /// </summary>
                /// <param name="sqlCmd">参数化命令</param>
                /// <param name="values">参数值列表</param>
                /// <returns>受影响的行数</returns>
                public virtual int ExecuteCommand(string sqlCmd, params object[] values)
                {
                        if (values != null && values.Length > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, values);
                                return Execute();
                        }
                        else
                        {
                                return ExecuteCommand(sqlCmd);
                        }
                }

                /// <summary>
                /// 执行参数化查询，如果命令中不带参数则抛出ArgumentException异常
                /// </summary>
                /// <param name="sqlCmd"></param>
                /// <param name="values"></param>
                /// <returns></returns>
                public virtual int ExecuteCommand(string sqlCmd, IDictionary<string, object> values)
                {
                        if (values != null && values.Count > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, values);
                                return Execute();
                        }
                        else
                        {
                                return ExecuteCommand(sqlCmd);
                        }
                }

                public virtual int Execute()
                {
                        return ExecuteCommand();
                }

                /// <summary>
                /// 执行SQL命令
                /// </summary>
                /// <returns>SQL命令影响行数</returns>
                public virtual int ExecuteCommand()
                {
                        int result = -1;
                        try
                        {
                                this.Command.Connection = this.Connection;
                                if (this.Command.Connection == null)
                                {
                                        throw new DALCoreException("no_connection");
                                }

                                if (this.Command.CommandText.Length > 0)
                                {
                                        this.Connection.Open();                                                     //开启数据库连接
                                        this.OpenSession(true);

                                        result = this.Command.ExecuteNonQuery();

                                        //if (this._transaction != null)
                                        //{
                                        //    this._transaction.Commit();
                                        //}
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
                                ResetCommand();
                                this.Connection.Close();      //关闭数据库连接
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
                                if (this.Command != null)
                                {
                                        this.Command.CommandText = sqlCmd;
                                        this.Command.Connection = this.Connection;

                                        this.OpenSession();

                                        return this.Command.ExecuteReader(CommandBehavior.CloseConnection);
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
                                ResetCommand();
                                this.CloseSession();
                        }
                }

                /// <summary>
                /// 根据命令对象得到一个数据记录游标
                /// </summary>
                /// <returns></returns>
                public virtual IDataReader GetDataReader()
                {
                        return GetDataReaderFromCommand();
                }

                /// <summary>
                /// 根据命令对象得到一个数据记录游标
                /// </summary>
                /// <returns></returns>
                public virtual IDataReader GetDataReaderFromCommand()
                {
                        try
                        {
                                this.Command.Connection = this.Connection;
                                if (this.Command.Connection == null)
                                {
                                        throw new DALCoreException("no_connection");
                                }

                                if (this.Command.CommandText.Length > 0)
                                {
                                        this.OpenSession();
                                        return this.Command.ExecuteReader(CommandBehavior.CloseConnection);
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
                                ResetCommand();
                                this.CloseSession();
                        }
                }

                public virtual IDataReader GetDataReaderFromCommand(string sqlCmd, params object[] paramValues)
                {
                        if (paramValues != null && paramValues.Length > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, paramValues);
                                return GetDataReaderFromCommand();
                        }
                        else
                        {
                                return GetDataReaderFromCommand(sqlCmd);
                        }
                }

                private void ResetCommand()
                {
                        Command.Parameters.Clear();
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
                                if (this.Command != null)
                                {
                                        this.Command.CommandText = sqlCmd;
                                        this.Command.Connection = this.Connection;

                                        this.OpenSession();

                                        return this.Command.ExecuteScalar();
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
                                ResetCommand();
                                this.CloseSession();
                        }
                }

                /// <summary>
                /// 取得结果集中第一行第一列的值，以对象类型返回
                /// </summary>
                /// <returns></returns>
                public virtual object GetScalarObject()
                {
                        return GetScalarObjectFromCommand();
                }

                /// <summary>
                /// 取得结果集中第一行第一列的值，以对象类型返回
                /// </summary>
                /// <returns></returns>
                public virtual object GetScalarObjectFromCommand()
                {
                        try
                        {
                                this.Command.Connection = this.Connection;
                                if (this.Command.Connection == null)
                                {
                                        throw new DALCoreException("no_connection");
                                }

                                if (this.Command.CommandText.Length > 0)
                                {
                                        this.OpenSession();

                                        return this.Command.ExecuteScalar();
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
                                ResetCommand();
                                this.CloseSession();
                        }
                }

                /// <summary>
                /// 参数化查询：取得结果集中第一行第一列的值，以对象类型返回
                /// </summary>
                /// <param name="sqlCmd">带参数的SQL</param>
                /// <param name="paramValues">参数值</param>
                /// <returns></returns>
                public virtual object GetScalarObjectFromCommand(string sqlCmd, params object[] paramValues)
                {
                        if (paramValues != null && paramValues.Length > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, paramValues);
                                return GetScalarObjectFromCommand();
                        }
                        else
                        {
                                return GetScalarObjectFromCommand(sqlCmd);
                        }
                }

                /// <summary>
                /// 获取数据集
                /// </summary>
                /// <param name="sqlCmd">SQL命令文本</param>
                /// <returns></returns>
                public virtual DataSet GetDataSetFromCommand(string sqlCmd)
                {
                        DataSet ds = new DataSet();
                        try
                        {
                                if (this.Command != null)
                                {
                                        this.Command.CommandText = sqlCmd;
                                        this.Command.Connection = this.Connection;

                                        if (this.DataAdapter != null)
                                        {
                                                this.DataAdapter.SelectCommand = this.Command;
                                                this.DataAdapter.Fill(ds);
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
                        finally
                        {
                                ResetCommand();
                        }

                        return ds;
                }

                /// <summary>
                /// 获取数据集
                /// </summary>
                /// <returns></returns>
                public virtual DataSet GetDataSet()
                {
                        return GetDataSetFromCommand();
                }

                /// <summary>
                /// 获取数据集
                /// </summary>
                /// <returns></returns>
                public virtual DataSet GetDataSetFromCommand()
                {
                        DataSet ds = new DataSet();

                        try
                        {
                                this.Command.Connection = this.Connection;
                                if (this.Command.Connection == null)
                                {
                                        throw new DALCoreException("no_connection");
                                }

                                if (this.Command.CommandText.Length > 0)
                                {
                                        if (this.DataAdapter != null)
                                        {
                                                this.DataAdapter.SelectCommand = this.Command;

                                                this.DataAdapter.Fill(ds);
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
                        finally
                        {
                                ResetCommand();
                        }

                        return ds;
                }

                /// <summary>
                /// 获取数据集
                /// </summary>
                /// <param name="sqlCmd">参数化查询</param>
                /// <param name="paramValues">参数值</param>
                /// <returns></returns>
                public virtual DataSet GetDataSetFromCommand(string sqlCmd, params object[] paramValues)
                {
                        if (paramValues != null && paramValues.Length > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, paramValues);
                                return GetDataSetFromCommand();
                        }
                        else
                        {
                                return GetDataSetFromCommand(sqlCmd);
                        }
                }

                /// <summary>
                /// 获取数据表
                /// </summary>
                /// <returns></returns>
                public virtual DataTable GetDataTable()
                {
                        return GetDataTableFromCommand();
                }

                /// <summary>
                /// 获取数据表
                /// </summary>
                /// <param name="command">命令对象</param>
                /// <returns></returns>
                public virtual DataTable GetDataTableFromCommand()
                {
                        DataTable dtResult = new DataTable();

                        try
                        {
                                DataSet dsTmp = this.GetDataSetFromCommand();
                                dtResult = dsTmp.Tables[0];
                                dsTmp.Tables.Clear();
                        }
                        catch (DALCoreException e)
                        {
                                throw e;
                        }
                        finally
                        {
                                ResetCommand();
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
                                if (this.Command != null)
                                {
                                        this.Command.CommandText = sqlCmd;
                                        this.Command.Connection = this.Connection;
                                        dtResult = this.GetDataTableFromCommand();
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
                        finally
                        {
                                ResetCommand();
                        }

                        return dtResult;
                }

                /// <summary>
                /// 获取数据表
                /// </summary>
                /// <param name="sqlCmd">参数化查询命令</param>
                /// <param name="paramValues">参数值</param>
                /// <returns></returns>
                public virtual DataTable GetDataTableFromCommand(string sqlCmd, params object[] paramValues)
                {
                        if (paramValues != null && paramValues.Length > 0)
                        {
                                PrepareParameterizedCommand(sqlCmd, paramValues);
                                return GetDataTableFromCommand();
                        }
                        else
                        {
                                return GetDataTableFromCommand(sqlCmd);
                        }
                }

                /// <summary>
                /// 开启数据库事务
                /// </summary>
                public void BeginTransaction()
                {
                        this._transaction = this.Connection.BeginTransaction();
                }

                /// <summary>
                /// 取得下一个序列
                /// </summary>
                /// <param name="tableName"></param>
                /// <returns></returns>
                public virtual string GetNextID(string tableName)
                {
                        return ConvertUtil.ToStringOrDefault(GetScalarObjectFromCommand(
                                string.Format(@"SELECT SEQ_{0}.nextval FROM dual", tableName)));
                }
        }
}