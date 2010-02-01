using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Principal;
using System.Collections;

namespace EasyDev.Security
{
    /// <summary>
    /// 用户身份识别对象
    /// </summary>
    public sealed class UserIdentity : GenericIdentity
    {
        private IDictionary context = null;

        //private bool isAuthenticated = true;

        //public void SetIsAuthenticated()
        //{
        //    this.isAuthenticated = true;
        //}

        /// <summary>
        /// 包含用户身份相关的信息集合
        /// </summary>
        public IDictionary Context
        {
            get
            {
                return this.context;
            }
            set
            {
                this.context = value;
            }
        }

        public UserIdentity(string userName)
            : base(userName)
        {            
            context = new Hashtable();
        }

        public UserIdentity(string userName, string type)
            : base(userName, type)
        {
            context = new Hashtable();
        }
    }
}
