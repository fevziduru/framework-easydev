using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace EasyDev.PL
{
        /// <summary>
        /// ADO对像扩展方法
        /// </summary>
        public static class ADOExtension
        {
                #region 私有方法

                private static string BuildValues(DataTable dt, DataRow dr)
                {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataColumn column in dt.Columns)
                        {
                                if (column.DataType == typeof(string) || column.DataType == typeof(Guid))
                                {
                                        sb.AppendFormat("'{0}',", dr[column.ColumnName] == DBNull.Value ? "" : dr[column.ColumnName]);
                                }
                                else if (column.DataType == typeof(DateTime))
                                {
                                        sb.AppendFormat("to_date('{0}','yyyy-mm-dd hh24:mi:ss'),", Convert.ToString(dr[column.ColumnName]));
                                }
                                else if ((column.DataType == typeof(decimal))
                                     || (column.DataType == typeof(int))
                                     || (column.DataType == typeof(double))
                                     || (column.DataType == typeof(float))
                                )
                                {
                                        sb.AppendFormat("{0},", dr[column.ColumnName] == DBNull.Value ? 0 : dr[column.ColumnName]);
                                }
                                else if (column.DataType == typeof(bool))
                                {
                                        sb.AppendFormat("'{0}',", dr[column.ColumnName] == DBNull.Value ? false : dr[column.ColumnName]);
                                }
                                else
                                {
                                        sb.AppendFormat("{0},", dr[column.ColumnName]);
                                }
                        }

                        if (sb.Length > 1)
                        {
                                sb = sb.Remove(sb.Length - 1, 1);
                        }
                        return sb.ToString();
                }

                private static string BuildSetClauseForUpdate(DataTable dt, DataRow dr)
                {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataColumn column in dt.Columns)
                        {
                                sb.Append(GetConditionByColumns(column, dr, ","));
                        }
                        string result = sb.ToString();
                        if (result.Length > 0)
                        {
                                result = result.Trim().TrimEnd(new char[] { ',' });
                        }

                        return result;
                }

                private static string GetConditionByColumns(DataColumn column, DataRow dr, string splitStr)
                {
                        string cond = "";
                        if (column.DataType == typeof(string)
                            || column.DataType == typeof(Guid))
                        {
                                cond = string.Format("{0}='{1}' {2} ", column.ColumnName, dr[column.ColumnName], splitStr);
                        }
                        else if (column.DataType == typeof(DateTime))
                        {
                                cond = string.Format("{0}=to_date('{1}','yyyy-mm-dd hh24:mi:ss') {2}", column.ColumnName, dr[column.ColumnName], splitStr);
                        }
                        else if ((column.DataType == typeof(decimal))
                                 || (column.DataType == typeof(int))
                                 || (column.DataType == typeof(double))
                                 || (column.DataType == typeof(float))
                            )
                        {
                                cond = string.Format("{0}={1} {2} ", column.ColumnName, dr[column.ColumnName], splitStr);
                        }

                        return cond;
                }

                private static string GetConditionByPrimaryKeys(string key, object value, string splitStr)
                {
                        string cond = "";
                        if (value.GetType() == typeof(string))
                        {
                                cond = string.Format(" {0}='{1}' {2}", key, value, splitStr);
                        }
                        else if ((value.GetType() == typeof(decimal))
                                 || (value.GetType() == typeof(int))
                                 || (value.GetType() == typeof(double))
                                 || (value.GetType() == typeof(float))
                            )
                        {
                                cond = string.Format(" {0}={1} {2}", key, value, splitStr);
                        }
                        else if (value.GetType() == typeof(DateTime))
                        {
                                cond = string.Format(" {0}='{1}' {2}", key, value, splitStr);
                        }

                        return cond;
                }

                /// <summary>
                /// 生成主键列
                /// </summary>
                /// <param name="dt">主键所在的表</param>
                /// <param name="dr">数据行</param>
                /// <returns></returns>
                private static string BuildPrimaryColumns(DataTable dt, DataRow dr)
                {
                        StringBuilder sb = new StringBuilder();
                        try
                        {
                                foreach (DataColumn column in dt.PrimaryKey)
                                {
                                        sb.Append(GetConditionByColumns(column, dr, " AND"));
                                }
                        }
                        catch (PersistenceException e)
                        {
                                throw e;
                        }

                        string result = sb.ToString();
                        if (result.Length > 0)
                        {
                                result = result.Trim().Remove(result.Length - " AND".Length - 1);
                        }

                        return result;
                }

                /// <summary>
                /// 生成主键列
                /// </summary>
                /// <param name="dr">数据行</param>
                /// <returns></returns>
                private static string BuildPrimaryColumns(DataRow dr)
                {
                        return BuildPrimaryColumns(dr.Table, dr);
                }
                #endregion

                /// <summary>
                /// 取得字段列表
                /// </summary>
                /// <param name="ds"></param>
                /// <returns></returns>
                public static string GetColumns(this DataTable dt)
                {
                        StringBuilder sb = new StringBuilder();

                        if (dt != null)
                        {
                                foreach (DataColumn column in dt.Columns)
                                {
                                        sb.AppendFormat("{0},", column.ColumnName);
                                }

                                //移除最后一个逗号
                                sb = sb.Remove(sb.Length - 1, 1);
                        }

                        return sb.ToString();
                }

                /// <summary>
                /// 创建INSERT语句
                /// </summary>
                /// <param name="fullName">INSERT语句操作的数据表名</param>
                /// <param name="dt"></param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string InsertCommandText(this DataTable dt, DataRow dataRow, string fullName)
                {
                        StringBuilder builder = new StringBuilder();

                        builder.AppendFormat(@"INSERT INTO {0} ({1}) VALUES ({2})", fullName, dt.GetColumns(), BuildValues(dt, dataRow));

                        return builder.ToString();
                }

                /// <summary>
                /// 创建INSERT语句
                /// </summary>
                /// <param name="fullName">INSERT语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string InsertCommandText(this DataRow dataRow, string fullName)
                {
                        return dataRow.Table.InsertCommandText(dataRow, fullName);
                }

                /// <summary>
                /// 创建DELETE语句
                /// </summary>
                /// <param name="fullName">DELETE语句操作的数据表名</param>
                /// <param name="dt"></param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string DeleteCommandText(this DataTable dt, DataRow dataRow,string fullName)
                {
                        StringBuilder builder = new StringBuilder();

                        try
                        {
                                DataColumn[] primaryKeys = dt.PrimaryKey;
                                if (primaryKeys.Length > 0)
                                {
                                        builder.AppendFormat(@"DELETE FROM {0} WHERE {1}", fullName, BuildPrimaryColumns(dt, dataRow));
                                }
                                else
                                {
                                        throw new PersistenceException("_no_primary_key");
                                }
                        }
                        catch (PersistenceException e)
                        {
                                throw e;
                        }

                        return string.Empty;
                }

                /// <summary>
                /// 创建DELETE语句
                /// </summary>
                /// <param name="fullName">DELETE语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string DeleteCommandText(this DataRow dataRow, string fullName)
                {
                        return dataRow.Table.DeleteCommandText(dataRow, fullName);
                }

                /// <summary>
                /// 创建UPDATE语句
                /// </summary>
                /// <param name="dt">UPDATE语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string UpdateCommandText(this DataTable dt, DataRow dataRow, string fullName)
                {
                        StringBuilder builder = new StringBuilder();

                        try
                        {
                                builder.AppendFormat(@"UPDATE {0} SET {1} WHERE {2}"
                                                     , fullName, BuildSetClauseForUpdate(dt, dataRow), BuildPrimaryColumns(dt, dataRow));
                        }
                        catch (PersistenceException e)
                        {
                                throw e;
                        }

                        return builder.ToString();
                }

                /// <summary>
                /// 创建UPDATE语句
                /// </summary>
                /// <param name="fullName">UPDATE语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string UpdateCommandText(this DataRow dataRow, string fullName)
                {
                        return dataRow.Table.UpdateCommandText(dataRow, fullName);
                }
        }
}
