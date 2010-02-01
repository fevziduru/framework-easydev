using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.PL
{
    /// <summary>
    /// Oracle序列值生成策略
    /// </summary>
    public class OracleSequenceGenerator : IGenerator
    {
        /// <summary>
        /// 序列对应的表名
        /// </summary>
        public string TableName { get; set; }

        /// <summary>
        /// 数据库会话对象
        /// </summary>
        public GenericDBSession Session { get; set; }

        /// <summary>
        /// 构造方法
        /// </summary>
        public OracleSequenceGenerator() { }

        #region IGenerator 成员

        public string Generate()
        {
            try
            {
                return Convert.ToString(
                    Session.GetScalarObjectFromCommand(
                    string.Format(@"SELECT SEQ_{0}.nextval FROM dual", TableName)));
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #endregion
    }
}
