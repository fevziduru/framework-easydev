using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.PL
{
    /// <summary>
    /// 普通主键生成器，生成GUID编码
    /// </summary>
    public class CommonIdentityGenerator : IGenerator
    {
        public CommonIdentityGenerator() { }

        #region IGenerator 成员

        /// <summary>
        /// 生成GUID编码作为主键值
        /// </summary>
        /// <returns></returns>
        public string Generate()
        {
            return Guid.NewGuid().ToString();
        }

        #endregion
    }
}
