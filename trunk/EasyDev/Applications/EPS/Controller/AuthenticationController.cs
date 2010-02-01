using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using WebSecurity = System.Web.Security;
using System.Web.SessionState;
using EasyDev.Security;
using System.Collections;
using System.Web.Security;

namespace EasyDev.EPS
{
    public class AuthenticationController : GenericController
    {
        protected IDictionary UserInfo
        {
            get;
            private set;
        }

        protected override void OnAuthorization(AuthorizationContext filterContext)
        {
            base.OnAuthorization(filterContext);
        }

        /// <summary>
        /// Action执行前处理
        /// </summary>
        /// <param name="filterContext"></param>
        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            object[] attributes = 
                filterContext.ActionDescriptor.GetCustomAttributes(typeof(EnsureLoginAttribute), true);
            
            try
            {
                //attributes不为NULL则需要用户验证
                if (attributes != null && attributes.Length > 0)
                {
                    #region 用户验证
                    if (Session != null && Session["USER_INFO"] != null)
                    {
                        UserInfo = Session["USER_INFO"] as IDictionary;
                    }
                    else
                    {
                        FormsAuthentication.RedirectToLoginPage();
                    }
                    #endregion
                }
                else
                {
                    UserInfo = null;
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            base.OnActionExecuting(filterContext);
        }
    }
}
