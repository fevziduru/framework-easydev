using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Collections;
using System.Data.Common;

namespace EasyDev.PL
{
        /// <summary>
        /// SQL代码生成器
        /// </summary>
        public class SqlBuilder
        {
                #region 私有方法

                private static string BuildValues(DataTable dt, DataRow dr)
                {
                        StringBuilder sb = new StringBuilder();

                        foreach (DataColumn column in dt.Columns)
                        {
                                if (column.DataType == typeof(string) || column.DataType == typeof(Guid))
                                {
                                        sb.AppendFormat("'{0}',", dr[column.ColumnName] == null ? "" : dr[column.ColumnName]);
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
                                        sb.AppendFormat("{0},", dr[column.ColumnName] == null ? 0 : dr[column.ColumnName]);
                                }
                                else if (column.DataType == typeof(bool))
                                {
                                        sb.AppendFormat("'{0}',", dr[column.ColumnName] == null ? false : dr[column.ColumnName]);
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

                public static string GetConditionByColumns(DataColumn column, DataRow dr, string splitStr)
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

                        //if (cond.Length > 0)
                        //{
                        //    cond = cond.Substring(0, cond.Length - 1);
                        //}
                        return cond;
                }

                public static string GetConditionByPrimaryKeys(string key, object value, string splitStr)
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

                        //if (cond.Length > 0)
                        //{
                        //    cond = cond.Substring(0, cond.Length - splitStr.Length - 1);
                        //}
                        return cond;
                }

                #endregion

                /// <summary>
                /// 生成主键列
                /// </summary>
                /// <param name="dt">主键所在的表</param>
                /// <param name="dr">数据行</param>
                /// <returns></returns>
                public static string BuildPrimaryColumns(DataTable dt, DataRow dr)
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
                public static string BuildPrimaryColumns(DataRow dr)
                {
                        return BuildPrimaryColumns(dr.Table, dr);
                }

                /// <summary>
                /// 取得字段列表
                /// </summary>
                /// <param name="dt"></param>
                /// <returns></returns>
                public static string GetColumns(DataTable dt)
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
                public static string BuildInsertCommand(string fullName, DataTable dt, DataRow dr)
                {
                        StringBuilder builder = new StringBuilder();

                        builder.AppendFormat
                            (@"INSERT INTO {0} ({1}) VALUES ({2})",
                             fullName, GetColumns(dt), BuildValues(dt, dr));

                        return builder.ToString();
                }

                /// <summary>
                /// 创建INSERT语句
                /// </summary>
                /// <param name="fullName">INSERT语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string BuildInsertCommand(string fullName, DataRow dr)
                {
                        return BuildInsertCommand(fullName, dr.Table, dr);
                }

                /// <summary>
                /// 创建DELETE语句
                /// </summary>
                /// <param name="fullName">DELETE语句操作的数据表名</param>
                /// <param name="dt"></param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string BuildDeleteCommand(string fullName, DataTable dt, DataRow dr)
                {
                        StringBuilder builder = new StringBuilder();

                        try
                        {
                                DataColumn[] primaryKeys = dt.PrimaryKey;
                                if (primaryKeys.Length > 0)
                                {
                                        builder.AppendFormat(@"DELETE FROM {0} WHERE {1}", fullName, BuildPrimaryColumns(dt, dr));
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
                public static string BuildDeleteCommand(string fullName, DataRow dr)
                {
                        return BuildDeleteCommand(fullName, dr.Table, dr);
                }

                /// <summary>
                /// 创建UPDATE语句
                /// </summary>
                /// <param name="dt">UPDATE语句操作的数据表名</param>
                /// <param name="dr"></param>
                /// <returns></returns>
                public static string BuildUpdateCommand(string fullName, DataTable dt, DataRow dr)
                {
                        StringBuilder builder = new StringBuilder();

                        try
                        {
                                builder.AppendFormat(@"UPDATE {0} SET {1} WHERE {2}"
                                                     , fullName, BuildSetClauseForUpdate(dt, dr), BuildPrimaryColumns(dt, dr));
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
                public static string BuildUpdateCommand(string fullName, DataRow dr)
                {
                        return BuildUpdateCommand(fullName, dr.Table, dr);
                }
        }
}
