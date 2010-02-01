using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Principal;
using System.Web;
using System.Collections;
using System.Threading;
using System.Web.SessionState;
using System.Web.Security;

namespace EasyDev.Security
{
    /// <summary>
    /// 用户安全机制管理器
    /// </summary>
    public class UserSecurityManager
    {
        /// <summary>
        /// 用户安全上下文
        /// </summary>
        public static IDictionary Context
        {
            get
            {
                IDictionary result = null;
                Principal principal = Thread.CurrentPrincipal as Principal;
                if (principal != null)
                {
                    UserIdentity identity = principal.Identity as UserIdentity;
                    if (identity != null)
                    {
                        result = identity.Context;
                    }
                }

                return result;
            }
        }

        /// <summary>
        /// 用户信息
        /// </summary>
        public static UserInfo UserInfo
        {
            get
            {
                UserInfo result = null;
                IDictionary context = Context;
                if (context != null)
                {
                    if (context["USER_INFO"] != null && context.Contains("USER_INFO"))
                    {
                        result = context["USER_INFO"] as UserInfo;
                    }
                }

                return result == null ? new UserInfo() : result;
            }
        }

        public static IDictionary CurrentUser
        {
            get
            {
                if (Context != null)
                {
                    if (Context["USER_INFO"] != null && Context.Contains("USER_INFO"))
                    {
                        return Context["USER_INFO"] as IDictionary;
                    }
                    else
                    {
                        return new Hashtable();
                    }
                }
                else
                {
                    return new Hashtable();
                }
            }
        }

        /// <summary>
        /// 将用户信息写入SESSION和当前线程
        /// </summary>
        public static void InitializeUserContext(HttpSessionState session, UserInfo user)
        {
            FormsAuthentication.SetAuthCookie(user.UserName, true);
            IDictionary items = new Hashtable();
            items.Add("USER_INFO",user);
            UserAuthentication.CheckIn(items);

            if (session != null)
            {
                if (session["USER_INFO"] != null)
                {
                    session["USER_INFO"] = user;
                }
                else
                {
                    session.Add("USER_INFO", user);
                }
            }
        }

        /// <summary>
        /// 移除用户信息
        /// </summary>
        /// <param name="session"></param>
        public static void RemoveUserContext(HttpSessionState session)
        {
            FormsAuthentication.SignOut();
            UserAuthentication.CheckOut();
            session.Remove("USER_INFO");
        }
    }
}
