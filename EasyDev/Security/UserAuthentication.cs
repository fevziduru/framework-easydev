using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Security.Principal;
using System.Threading;

namespace EasyDev.Security
{
    /// <summary>
    /// 用户验证对象
    /// </summary>
    public class UserAuthentication
    {
        #region IUserAuthentication Members

        /// <summary>
        /// 用户签入
        /// </summary>
        /// <param name="context"></param>
        public static void CheckIn(IDictionary context)
        {
            UserIdentity identity = new UserIdentity(string.Empty);
            //identity.SetIsAuthenticated();
            identity.Context = context;
            Thread.CurrentPrincipal = new Principal(identity);
        }

        /// <summary>
        /// 用户签出
        /// </summary>
        public static void CheckOut()
        {
            UserIdentity identity = new UserIdentity(string.Empty);
            Thread.CurrentPrincipal = null;
        }

        /// <summary>
        /// 验证某用户对某资源是否有某种操作类型的权限
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="resourceId">资源ID</param>
        /// <param name="operationId">操作ID</param>
        /// <returns></returns>
        public static bool Authentication(string userName, string resourceId, string operationId)
        {
            //TODO:验证某用户对某资源是否有某种操作类型的权限

            return true;
        }

        /// <summary>
        /// 验证某用户对某资源是否有任一操作类型的权限
        /// </summary>
        /// <param name="userName">用户名</param>
        /// <param name="resourceId">资源ID</param>
        /// <returns></returns>
        public static bool Authentication(string userName, string resourceId)
        {
            //TODO:验证某用户对某资源是否有任一操作类型的权限

            return true;
        }

        #endregion
    }
}
