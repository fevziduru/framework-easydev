using System;
using System.Data;
using EasyDev.PL;
using System.Text;
using System.Collections.Generic;
using System.Collections;

namespace EasyDev.BL
{
    /// <summary>
    /// 通用业务对象，用于执行常用的数据存取操作，通过内部的数据库会话对象与数据库进行交互，并将数据保存在数据集（强类型或弱类型）中。
    /// 另外，此对象也支持从数据库中加载视图数据，通过视图名称将通过定义视图的查询加载数据到视图数据集合中
    /// </summary>
    /// <remarks>此对象可被序列化</remarks>
    /// <typeparam name="T">数据集类型，此类型必须由DataSet类继承</typeparam>
    [Serializable]
    public class GenericBizObject<T> : IGenericBizObject, IDisposable
        where T : DataSet
    {
        #region 私有成员
        private T _data = null;
        private DataSet views = null;
        private string _schema = string.Empty;
        private GenericDBSession _session = null;
        private string _entityName = string.Empty;
        #endregion

        #region 公共属性
        /// <summary>
        /// 通用业务对象的名称，在服务中初始化时与服务对应的xsd数据集架构文件同名
        /// </summary>
        public string BOName
        {
            get;
            set;
        }

        /// <summary>
        /// 数据表全名
        /// </summary>
        public string FullName { get; set; }

        /// <summary>
        /// 业务对象表示的表名称
        /// </summary>
        public string Entity
        {
            get { return this._entityName; }
            set { this._entityName = value; }
        }

        /// <summary>
        /// 数据集架构，对应配置信息中的xsd数据集架构文件名称
        /// </summary>
        public string Schema
        {
            get { return this._schema; }
            set { this._schema = value; }
        }

        /// <summary>
        /// 与业务对象关联的数据库会话
        /// </summary>
        public GenericDBSession CurrentSession
        {
            get
            {
                return this._session;
            }
        }

        /// <summary>
        /// 业务对象数据
        /// </summary>
        public T Data
        {
            get { return _data; }
            set { _data = value; }
        }

        /// <summary>
        /// 业务对象所包含的数据
        /// </summary>
        public DataSet BOData
        {
            get { return (DataSet)_data; }
            set { _data = (T)value; }
        }

        /// <summary>
        /// 视图集合数据
        /// </summary>
        public DataSet Views
        {
            get
            {
                return this.views;
            }
        }
        #endregion

        #region 公共构造方法

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="ds"><seealso cref="GenericBizObject<T>">T类型</seealso>数据集</param>
        public GenericBizObject(T ds)
        {
            this._data = ds;
            this.views = new DataSet();
        }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="ds"><seealso cref="GenericBizObject<T>"/></param>
        /// <param name="session">与当前业务对象关联的数据库会话对象</param>
        public GenericBizObject(T ds, GenericDBSession session, string tableName)
        {
            this._data = ds;
            this._session = session;
            this.BOName = tableName;
            //if (CurrentSession.CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
            //    CurrentSession.CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
            //{
            //    ((OracleSequenceGenerator)CurrentSession.IdentityGenerator).TableName = tableName;
            //}

            this.views = new DataSet();
        }
        #endregion

        #region IDisposable 成员

        public void Dispose()
        {
            this._data = null;
            this.views = null;
        }

        #endregion

        #region IBizObject 成员

        /// <summary>
        /// 根据行状态将数据同步到数据库
        /// </summary>
        /// <exception cref="PersistenceException">数据同步出错时将引发异常</exception>
        /// <remarks>
        /// 数据同步方法只同步与当前业务对象相关的数据，即当前BOName对应的xsd数据集架构文件相关的数据表
        /// </remarks>
        public virtual void Synchronize()
        {
            try
            {
                if (this.BOData.Tables.Count > 0 && this.BOData.Tables.Contains(BOName))
                {
                    DataTable dt = this.BOData.Tables[BOName];
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row.RowState == DataRowState.Added)
                        {
                            //_session.ExecuteCommand(SqlBuilder.BuildInsertCommand(this.FullName, row.Table, row));

                            //TODO: ADO扩展方法测试
                            _session.ExecuteCommand(row.Table.InsertCommandText(row, this.FullName));
                        }
                        else if (row.RowState == DataRowState.Modified)
                        {
                            //_session.ExecuteCommand(SqlBuilder.BuildUpdateCommand(this.FullName, row.Table, row));

                            //TODO: ADO扩展方法测试
                            _session.ExecuteCommand(row.Table.UpdateCommandText(row, FullName));
                        }
                        else if (row.RowState == DataRowState.Deleted)
                        {
                            //_session.ExecuteCommand(SqlBuilder.BuildDeleteCommand(this.FullName, row.Table, row));

                            //TODO: ADO扩展方法测试
                            _session.ExecuteCommand(row.Table.DeleteCommandText(row, FullName));
                        }
                    }

                    //this._data.Tables[0].AcceptChanges();
                    dt.AcceptChanges();
                }
            }
            catch (PersistenceException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 加载所有数据
        /// </summary>
        /// <returns>返回true则成功</returns>
        /// <remarks>
        /// 此方法的返回值可以不用接收，因为此方法出错会抛出异常
        /// </remarks>
        /// <exception cref="PersistenceException">加载数据出错时抛出异常</exception>
        public virtual bool LoadAll()
        {
            return this.LoadByCondition("");
        }

        /// <summary>
        /// 根据主键加载数据
        /// </summary>
        /// <remarks>
        /// 此方法的返回值可以不用接收，因为此方法出错会抛出异常<br/>
        /// 此方法通过主键加载数据，而主键信息则包含在xsd数据集架构文件中
        /// </remarks>
        /// <param name="keyValue">主键值</param>
        /// <exception cref="PersistenceException">加载数据出错时抛出异常</exception>
        /// <returns>返回true则成功</returns>
        public virtual bool LoadByPrimaryKey(object keyValue)
        {
            string cond = "";
            try
            {
                DataColumn[] cols = this.BOData.Tables[this._entityName].PrimaryKey;
                string firstKey = "";
                if (cols.Length > 0)
                {
                    firstKey = cols[0].ColumnName;
                    DataColumn column = cols[0];
                    cond += SqlBuilder.GetConditionByPrimaryKeys(firstKey, keyValue, ",");

                    //如果条件串以逗号结尾则去掉它
                    if (cond.Trim().EndsWith(","))
                    {
                        cond = cond.TrimEnd(new char[] { ',' });
                    }
                }
            }
            catch (PersistenceException ex)
            {
                throw ex;
            }

            return this.LoadByCondition(cond);
        }

        /// <summary>
        /// 根据主键加载数据（多主键情况）
        /// </summary>
        /// <param name="keyValues">多主键的主属性名称及其值组成的键值对集合</param>
        /// <remarks>
        /// 此方法的返回值可以不用接收，因为此方法出错会抛出异常<br/>
        /// 此方法通过主键加载数据，而主键信息则包含在xsd数据集架构文件中<br/>
        /// 因为无法判断多主键情况下，多个主属性的排列顺序，所以通过IDictionary对象来传递多主键的值
        /// </remarks>
        /// <exception cref="PersistenceException">加载数据出错时抛出异常</exception>
        /// <returns>返回true时成功</returns>
        public virtual bool LoadByPrimaryKeys(IDictionary<string, object> keyValues)
        {
            string cond = "";
            try
            {
                //TODO: 测试
                DataColumn[] cols = this.BOData.Tables[this._entityName].PrimaryKey;
                if (cols.Length > 0 && keyValues != null)
                {
                    IEnumerator itr_keyvalues = keyValues.GetEnumerator();
                    while (itr_keyvalues.MoveNext())
                    {
                        KeyValuePair<string, object> item = (KeyValuePair<string, object>)itr_keyvalues.Current;
                        cond += SqlBuilder.GetConditionByPrimaryKeys(item.Key, item.Value, "AND");
                    }

                    if (cond.Trim().EndsWith("AND"))
                    {
                        cond = cond.Remove(cond.Length - "AND".Length);
                    }
                }
            }
            catch (PersistenceException ex)
            {
                throw ex;
            }

            return this.LoadByCondition(cond);
        }

        /// <summary>
        /// 根据条件加载数据
        /// </summary>
        /// <param name="whereCond">条件表达式</param>
        /// <exception cref="PersistenceException">加载数据出错时抛出异常</exception>
        /// <returns>返回true时成功</returns>
        public virtual bool LoadByCondition(string whereCond)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                if (string.IsNullOrEmpty(whereCond) == false)
                {
                    whereCond = " WHERE " + whereCond;
                }

                sb.AppendFormat(@"SELECT {0} FROM {1} {2}",
                    /*SqlBuilder.GetColumns(this._data.Tables[this._entityName])*/
                                BOData.Tables[Entity].GetColumns(), this.FullName, whereCond);         //TODO: ADO扩展方法测试

                #region 重新加载业务对象数据
                //对于有自关联的表，要先删除其自关联关系，否则无法移除表
                this.BOData.Relations.Clear();

                //清除数据集中的数据表
                this.BOData.Tables.Clear();

                //读取架构以重新获得与数据表相关的约束信息
                //@注 : 如果这里没有读取架构信息而直接将查询后得到的表或数据集赋值给数据集或数据集中的表会导致之前的架构信息丢失，
                //以至不能从数据集或数据表中取得主键、外键等相关信息
                this.BOData.ReadXmlSchema(this._schema);

                //将从数据库查询到的数据表合并到数据集
                this.BOData.Tables[this._entityName].Merge(
                    this._session.GetDataTableFromCommand(sb.ToString()), true, MissingSchemaAction.AddWithKey);

                //修改数据集中的表名，使其与XSD数据集架构文件同名以方便使用
                if (this.BOData.Tables.Count > 0)
                {
                    this.BOData.Tables[0].TableName = this._entityName;
                }
                #endregion
            }
            catch (PersistenceException e)
            {
                throw e;
            }

            return true;
        }

            /// <summary>
            /// 延迟实例化主键生成器，如果在构造函数中实例化会出现覆盖的现象
            /// </summary>
        private void EnsureIdentityGenerator()
        {
                if (CurrentSession.CurrentDbProvider.Equals("System.Data.OracleClient", StringComparison.CurrentCultureIgnoreCase) ||
               CurrentSession.CurrentDbProvider.Equals("Oracle.DataAccess.Client", StringComparison.CurrentCultureIgnoreCase))
                {
                        ((OracleSequenceGenerator)CurrentSession.IdentityGenerator).TableName = this.BOName;
                }
        }

        /// <summary>
        /// 创建新行
        /// </summary>
        /// <remarks>
        /// 在创建新行的时候会同时处理不能为空的列，对于这些列会自动添加GUID值进行填充，在使用的时候可将其覆盖
        /// </remarks>
        /// <exception cref="PersistenceException">加载数据出错时抛出异常</exception>
        /// <returns>返回当前创建的新行</returns>
        public virtual DataRow Create()
        {
            DataRow drNew = null;
            try
            {
                    EnsureIdentityGenerator();

                if (this.BOData.Tables.Count > 0)
                {
                    DataTable dt = this.BOData.Tables[0];
                    drNew = dt.NewRow();
                    //处理主键
                    for (int i = 0; i < dt.PrimaryKey.Length; i++)
                    {
                        drNew[dt.PrimaryKey[i].ColumnName] = CurrentSession.IdentityGenerator.Generate();
                    }

                    //处理空值约束
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (dt.Columns[i].AllowDBNull == false)
                        {
                            if (dt.Columns[i].DataType == typeof(string))
                            {

                                drNew[i] = Guid.NewGuid().ToString().Substring(0,
                                        Guid.NewGuid().ToString().Length > dt.Columns[i].MaxLength ? dt.Columns[i].MaxLength : Guid.NewGuid().ToString().Length);
                            }
                            else if (dt.Columns[i].DataType == typeof(Guid))
                            {
                                drNew[i] = Guid.NewGuid();
                            }
                            else if (dt.Columns[i].DataType == typeof(decimal) || dt.Columns[i].DataType == typeof(double) || dt.Columns[i].DataType == typeof(int))
                            {
                                drNew[i] = 0;
                            }
                            else if (dt.Columns[i].DataType == typeof(DateTime))
                            {
                                drNew[i] = DateTime.Now;
                            }
                            else if (dt.Columns[i].DataType == typeof(bool))
                            {
                                drNew[i] = false;
                            }
                        }
                    }
                    dt.Rows.Add(drNew);
                }
            }
            catch (PersistenceException e)
            {
                throw e;
            }

            return drNew;
        }

        /// <summary>
        /// 判断业务对象数据是否为空
        /// </summary>
        /// <returns>返回true则数据为空</returns>
        [Obsolete("此方法已经过时，请使用IsEmpty方法")]
        public virtual bool IsModelEmpty()
        {
            return IsEmpty();
        }

        /// <summary>
        /// 判断业务对象数据是否为空
        /// </summary>
        /// <returns>返回true则数据为空</returns>
        public virtual bool IsEmpty()
        {
            return this.Data.Tables.Count <= 0
                || (this.Data.Tables.Count > 0 && this.Data.Tables[0].Rows.Count <= 0);
        }

        /// <summary>
        /// 取得业务对象数据中的第一行
        /// </summary>
        /// <remarks>
        /// 
        /// </remarks>
        /// <returns>返回业务对象数据中的第一行</returns>
        public virtual DataRow GetFirstRow()
        {
            if (IsEmpty() == false)
            {
                return this.Data.Tables[0].Rows[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 取得业务对象数据中的第一张表
        /// </summary>
        /// <returns>业务对象数据中的第一个数据表</returns>
        public virtual DataTable GetFirstTable()
        {
            if (IsEmpty() == false)
            {
                return this.Data.Tables[0];
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 清除业务对象数据
        /// </summary>
        public virtual void Clear()
        {
            this.Data.Clear();
        }

        #endregion

        /// <summary>
        /// 根据名称打开视图
        /// </summary>
        /// <param name="viewName">视图名称</param>
        /// <returns>视图数据</returns>
        public DataSet LoadViewByName(string viewName)
        {
            try
            {
                return LoadViewByCondition(viewName, null);
            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 根据条件打开视图
        /// </summary>
        /// <param name="viewName">视图名称</param>
        /// <param name="cond">条件表达式</param>
        /// <returns>视图数据</returns>
        public DataSet LoadViewByCondition(string viewName, string cond)
        {
            try
            {
                string sql = string.Empty;
                if (string.IsNullOrEmpty(cond))
                {
                    sql = string.Format(@"select * from {0}", viewName);
                }
                else
                {
                    sql = string.Format(@"select * from {0} where {1}", viewName, cond);
                }

                DataTable dtTmp = this._session.GetDataTableFromCommand(sql);

                dtTmp.TableName = viewName;
                if (this.views.Tables.Contains(viewName))
                {
                    this.views.Tables.Remove(viewName);
                }

                this.views.Tables.Add(dtTmp);
                return dtTmp.DataSet;
            }
            catch (System.Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 从视图集合中移除视图
        /// </summary>
        /// <param name="viewName">视图名称</param>
        public void RemoveFromViews(string viewName)
        {
            this.views.Tables.Remove(viewName);
        }

        /// <summary>
        /// 根据主键值删除数据
        /// </summary>
        /// <param name="key">主键值</param>
        /// <exception cref="PersistenceException">删除出错时抛出异常</exception>
        public void DeleteByPrimaryKey(object key)
        {
            string cond = string.Empty;
            try
            {
                DataColumn[] cols = this.BOData.Tables[this._entityName].PrimaryKey;
                string firstKey = "";

                if (cols.Length > 0)
                {
                    firstKey = cols[0].ColumnName;
                    DataColumn column = cols[0];
                    cond += SqlBuilder.GetConditionByPrimaryKeys(firstKey, key, ",");

                    //如果条件串以逗号结尾则去掉它
                    if (cond.Trim().EndsWith(","))
                    {
                        cond = cond.TrimEnd(new char[] { ',' });
                    }
                }

                this.DeleteByCondition(cond);
            }
            catch (PersistenceException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 多主键删除
        /// </summary>
        /// <param name="keys">主键键值对集合</param>
        /// <exception cref="PersistenceException">删除出错时抛出异常</exception>
        public void DeleteByPrimaryKeys(IDictionary<string, object> keys)
        {
            string cond = string.Empty;
            try
            {
                DataColumn[] cols = this.BOData.Tables[this._entityName].PrimaryKey;
                if (cols.Length > 0)
                {
                    IEnumerator itr_keyvalues = keys.GetEnumerator();
                    while (itr_keyvalues.MoveNext())
                    {
                        KeyValuePair<string, object> item = (KeyValuePair<string, object>)itr_keyvalues.Current;
                        cond += SqlBuilder.GetConditionByPrimaryKeys(item.Key, item.Value, "AND");
                    }

                    if (cond.Trim().EndsWith("AND"))
                    {
                        cond = cond.Remove(cond.Length - "AND".Length);
                    }
                }

                this.DeleteByCondition(cond);
            }
            catch (PersistenceException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 按条件删除
        /// </summary>
        /// <param name="cond">删除条件表达式</param>
        /// <exception cref="PersistenceException">删除出错时抛出异常</exception>
        public void DeleteByCondition(string cond)
        {
            try
            {
                StringBuilder deleteCommand = new StringBuilder();
                deleteCommand.Append(string.Format("DELETE FROM {0} WHERE {1}", this.Entity, cond));

                CurrentSession.ExecuteCommand(deleteCommand.ToString());
            }
            catch (PersistenceException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 根据表名取得对应的最新序列码
        /// </summary>
        /// <param name="tableName">表名</param>
        /// <remarks>
        /// 在数据库中要有以SEQ_[TABLENAME]命名的序列，否则会出现异常
        /// </remarks>        
        /// <exception cref="PersistenceException">如果没有找到序列则抛出异常</exception>
        /// <returns>序列值</returns>
        public string GetNextSequenceID(string tableName)
        {
            try
            {
                return Convert.ToString(
                    _session.GetScalarObjectFromCommand(
                    string.Format(@"SELECT SEQ_{0}.nextval FROM dual", tableName)));
            }
            catch (PersistenceException e)
            {
                throw e;
            }
        }
    }
}
