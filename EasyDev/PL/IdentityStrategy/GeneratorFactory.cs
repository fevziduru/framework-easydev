using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.PL
{
    /// <summary>
    /// 主键值生成器工厂
    /// </summary>
    public class GeneratorFactory
    {
        /// <summary>
        /// 创建主键值生成器
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public static T CreateGenerator<T>() 
            where T: IGenerator, new ()
        {
            return new T();
        }
    }
}
