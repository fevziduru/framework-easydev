using System;
using System.Linq.Expressions;
using System.Data;

namespace EasyDev.BL
{
        /// <summary>
        /// 
        /// </summary>
        public static class GenericServiceExtension
        {
                /// <summary>
                /// 
                /// </summary>
                /// <param name="srv"></param>
                /// <param name="expr"></param>
                /// <returns></returns>
                [Obsolete("此方法暂不使用", true)]
                public static DataTable Get(this GenericService srv, Expression<Func<DataRow, bool>> expr)
                {
                        DataSet ds = srv.LoadAll();
                        DataTable dt = ds.Tables[srv.BOName].Copy();
                        dt.Rows.Clear();

                        foreach (DataRow row in ds.Tables[srv.BOName].AsEnumerable().Where<DataRow>(expr.Compile()))
                        {
                                dt.ImportRow(row);
                        }

                        return dt;
                }
        }
}
