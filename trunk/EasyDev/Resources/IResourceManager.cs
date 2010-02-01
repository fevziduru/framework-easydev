using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Configuration;

namespace EasyDev.Resources
{
    public interface IResourceManager
    {
        /// <summary>
        /// 取得资源串值
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        string GetString(string key);
    }
}
