using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
    public class UserLoginService : AbstractService 
    {
        private UserService service = ServiceFactory.CreateService<UserService>();

        /// <summary>
        /// 登录方法
        /// </summary>
        /// <param name="username"></param>
        /// <param name="passwd"></param>
        /// <returns></returns>
        public virtual Users Login(string username, string passwd)
        {
            return service.ValidateUser(username, passwd);
        }

        public virtual void Logout()
        { 
            
        }
    }
}
