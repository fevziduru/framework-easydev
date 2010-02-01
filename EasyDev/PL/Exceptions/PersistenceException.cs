using EasyDev.Common;

namespace EasyDev.PL
{
    /// <summary>
    /// 持久层异常对象
    /// </summary>
    public class PersistenceException : ExceptionBase, IException
    {
        public PersistenceException(string msg)
            : base(msg)
        {}
    }
}
