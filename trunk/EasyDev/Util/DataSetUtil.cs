using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace EasyDev.Util
{
        [Obsolete("此工具类中的方法都已经有对应的扩展方法实现")]       
        public class DataSetUtil
        {
                /// <summary>
                /// 获取数据集中的数据表
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <returns></returns>
                public static DataTable GetDataTableFromDataSet(DataSet dsFrom, string tableName)
                {
                        return dsFrom.Table(tableName);
                }

                /// <summary>
                /// 获取数据集中的数据表
                /// </summary>
                /// <param name="dsFrom">源数据集</param>
                /// <param name="index">数据表索引</param>
                /// <returns></returns>
                public static DataTable GetDataTableFromDataSet(DataSet dsFrom, int index)
                {
                        return dsFrom.Table(index);
                }

                /// <summary>
                /// 获取数据集中某数据表的第一行
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <returns></returns>
                public static DataRow GetFirstRowFromDataSet(DataSet dsFrom, string tableName)
                {
                        return dsFrom.FirstRow(tableName);
                }

                /// <summary>
                /// 获取数据表中的第一行
                /// </summary>
                /// <param name="dtFrom">数据表</param>
                /// <returns></returns>
                public static DataRow GetFirstRowFromDataTable(DataTable dtFrom)
                {
                        return dtFrom.SingleRow(0);
                }

                /// <summary>
                /// 获取数据集中的一行
                /// </summary>
                /// <param name="dsFrom">数据集</param>
                /// <param name="tableName">数据表名称</param>
                /// <param name="rowNum">行号</param>
                /// <returns></returns>
                public static DataRow GetSingleRowFromDataSet(DataSet dsFrom, string tableName, int rowNum)
                {
                        return dsFrom.SingleRow(tableName, rowNum);
                }

                /// <summary>
                /// 获取数据表中的一行
                /// </summary>
                /// <param name="dtFrom">数据表</param>
                /// <param name="rowNum">行号</param>
                /// <returns></returns>
                public static DataRow GetSingleRowFromDataTable(DataTable dtFrom, int rowNum)
                {
                        return dtFrom.SingleRow(rowNum);
                }
        }
}
