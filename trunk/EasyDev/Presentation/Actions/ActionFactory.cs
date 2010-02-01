using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.Presentation
{
    /// <summary>
    /// 行为工厂
    /// </summary>
    public class ActionFactory
    {
        private ActionFactory() { }

        public static ActionFactory CreateInstance()
        {
            return new ActionFactory();
        }

        /// <summary>
        /// 创建行为
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T CreateAction<T>() where T : IAction, new()
        {
            return new T();
        }
    }
}
