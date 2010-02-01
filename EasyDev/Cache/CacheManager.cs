using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using WebCaching = System.Web.Caching;

namespace EasyDev.Cache
{
    public class CacheManager
    {
        private static WebCaching.Cache Container = HttpRuntime.Cache;

        static CacheManager()
        {
            Container = HttpRuntime.Cache;            
        }

        /// <summary>
        /// 将对象存入缓存
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public static void Store(string key, object value)
        {
            object tmp = Container.Get(key);
            if (tmp != null)
            {
                if (tmp.Equals(value) == false)
                {
                    Container.Insert(key, value);
                }
            }
            else
            {
                Container.Insert(key, value);
            }
        }

        /// <summary>
        /// 从缓存中取对象
        /// </summary>
        /// <param name="key"></param>
        /// <returns></returns>
        public static object Retrieve(string key)
        {
            return Container.Get(key);
        }
    }
}
