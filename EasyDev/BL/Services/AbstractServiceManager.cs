using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.BL
{
        /// <summary>
        /// 抽象服务管理器，用于实现默认的操作实现
        /// </summary>
        public class AbstractServiceManager : IServiceManager
        {
                public AbstractServiceManager() 
                {
                        Initialize();
                }
                
                protected virtual void Initialize() { }
        }
}
