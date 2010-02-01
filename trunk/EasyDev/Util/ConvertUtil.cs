using System;
using System.Collections.Generic;
using System.Text;
using System.Data;

namespace EasyDev.Util
{
        /// <summary>
        /// 
        /// </summary>
        [Obsolete("此工具类中的方法都已经有对应的扩展方法实现")]        
        public sealed class ConvertUtil
        {
                [Obsolete("此方法已经过时，请使用ToStringOrDefault")]
                public static string EmptyOrString(object obj)
                {
                        try
                        {
                                if (null == obj)
                                {
                                        return string.Empty;
                                }
                                else
                                {
                                        return System.Convert.ToString(obj);
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 转换为字符串，如果为空则返回string.Empty
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                [Obsolete("此方法已过时")]
                public static string ConvertToString(object obj)
                {
                        try
                        {
                                if (null == obj)
                                {
                                        return string.Empty;
                                }
                                else
                                {
                                        return System.Convert.ToString(obj);
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 返回对象的字符串表示,去掉前后空格
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static string ToStringOrDefault(object obj)
                {
                        return obj.ParseString();
                }

                /// <summary>
                /// 获取数据集中的数据表
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataTable GetDataTableFromDataSet(DataSet dsFrom, string tableName)
                {
                        try
                        {
                                if (dsFrom.Tables.Contains(tableName))
                                {
                                        return dsFrom.Tables[tableName];
                                }
                                else
                                {
                                        return null;
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 获取数据集中的数据表
                /// </summary>
                /// <param name="dsFrom">源数据集</param>
                /// <param name="index">数据表索引</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataTable GetDataTableFromDataSet(DataSet dsFrom, int index)
                {
                        try
                        {
                                if (dsFrom.Tables[index] != null)
                                {
                                        return dsFrom.Tables[index];
                                }
                                else
                                {
                                        return null;
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 获取数据集中某数据表的第一行
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataRow GetFirstRowFromDataSet(DataSet dsFrom, string tableName)
                {
                        try
                        {
                                return GetSingleRowFromDataSet(dsFrom, tableName, 0);
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 获取数据表中的第一行
                /// </summary>
                /// <param name="dtFrom">数据表</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataRow GetFirstRowFromDataTable(DataTable dtFrom)
                {
                        try
                        {
                                return GetSingleRowFromDataTable(dtFrom, 0);
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 获取数据集中的一行
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <param name="rowNum">行号</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataRow GetSingleRowFromDataSet(DataSet dsFrom, string tableName, int rowNum)
                {
                        try
                        {
                                if (dsFrom.Tables.Contains(tableName) && dsFrom.Tables[tableName].Rows.Count > 0)
                                {
                                        return dsFrom.Tables[tableName].Rows[rowNum];
                                }
                                else
                                {
                                        return null;
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 获取数据表中的一行
                /// </summary>
                /// <param name="dtFrom">数据表</param>
                /// <param name="rowNum">行号</param>
                /// <returns></returns>
                [Obsolete("请使用DataSetUtil中的对应方法")]
                public static DataRow GetSingleRowFromDataTable(DataTable dtFrom, int rowNum)
                {
                        try
                        {
                                if (dtFrom.Rows.Count > 0)
                                {
                                        return dtFrom.Rows[rowNum];
                                }
                                else
                                {
                                        return null;
                                }
                        }
                        catch (ConvertException e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 将对象转换成DateTime，如果无法转换则返回值为0001-1-1 00:00:00的DateTime（new DateTime();）
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static DateTime ToDateTime(object obj)
                {
                        return obj.ParseDateTime();
                }
                /// <summary>
                /// 将对象转换成decimal，如果无法转换则返回0
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static decimal ToDecimal(object obj)
                {
                        return obj.ParseDecimal();
                }
                /// <summary>
                /// 将对象转换成int，如果无法转换则返回0
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static int ToInt(object obj)
                {
                        return obj.ParseInt();
                }
                /// <summary>
                /// 转换成纬度，正数为北纬，负数为南纬，默认值为0度
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static decimal ToLat(object obj)
                {
                        return obj.ParseLatitude();
                }
                /// <summary>
                /// 转换成经度，正数为东经，负数为西经，默认值为0度
                /// </summary>
                /// <param name="obj"></param>
                /// <returns></returns>
                public static decimal ToLng(object obj)
                {
                        return obj.ParseLongitude();
                }
                /// <summary>
                /// 转换成bool类型，特别的，"Y"转换成True，"N"转换成False
                /// </summary>
                /// <param name="p"></param>
                /// <returns></returns>
                public static bool ToBool(object p)
                {
                        string str = ToStringOrDefault(p);
                        if ("N".Equals(str))
                        {
                                return false;
                        }
                        if ("Y".Equals(str))
                        {
                                return true;
                        }
                        try
                        {
                                return Convert.ToBoolean(p);
                        }
                        catch { }
                        return false;
                }

                public static TimeSpan ToTimeSpan(object obj, TimeUnit unit)
                {
                        //decimal num = Convert.ToDecimal(obj);
                        //if (num == 0.0M)
                        //{
                        //    return new TimeSpan();
                        //}
                        //decimal[] res1 = ResolveReal(num);

                        //TimeSpan s = new TimeSpan();
                        //switch (unit)
                        //{
                        //    case TimeUnit.Hour:
                        //        s.Add(new TimeSpan((int)res1[0],0,0);
                        //        decimal minutes = 60 * res1[1];
                        //        decimal[] res2 = ResolveReal(minutes);
                        //        s.Add(new TimeSpan(0,(int)res2[0],(in)))
                        //        return new TimeSpan(Convert.ToInt32(obj), 0, 0);
                        //    case TimeUnit.Minute:

                        //}
                        throw new NotImplementedException();
                }

                /// <summary>
                /// 
                /// </summary>
                /// <param name="d"></param>
                /// <returns></returns>
                public static decimal[] ResolveReal(decimal d)
                {
                        return d.ResolveReal();
                }
        }
}
