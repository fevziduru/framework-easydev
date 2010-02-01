using EasyDev.Common;

namespace EasyDev.PL
{
    /// <summary>
    /// 数据访问层异常对象
    /// </summary>
    public class DALCoreException : ExceptionBase, IException
    {
        public DALCoreException(string msg)
            : base(msg)
        { }
    }
}
