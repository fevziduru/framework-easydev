using System;

namespace EasyDev.PL
{
    /// <summary>
    /// 生成主键值的方法
    /// </summary>
    [Obsolete("此类已经过时，使用实现了IGenerator接口并通过GeneratorFactory创建的类")]
    public class IdentityGenerator
    {
        /// <summary>
        /// 生成GUID作为实体标识
        /// </summary>
        /// <returns></returns>
        public static string GUIDIdentity()
        {
            return Guid.NewGuid().ToString();
        }

        /// <summary>
        /// 生成ORACLE数据表对应的SEQUENCE
        /// </summary>
        /// <param name="session"></param>
        /// <param name="tableName"></param>
        /// <returns></returns>
        public static string GetNextSequenceId(GenericDBSession session, string tableName)
        {
            try
            {
                return Convert.ToString(
                    session.GetScalarObjectFromCommand(
                    string.Format(@"SELECT SEQ_{0}.nextval FROM dual", tableName)));
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
