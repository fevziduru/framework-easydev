using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Security;
using System.Security.Principal;
using System.Data;
using EasyDev.Util;
using System.Threading;
using System.Text.RegularExpressions;
using System.Configuration;
using log4net;
using EasyDev.BL;

namespace EasyDev.SQMS
{
        public class PermissionModule : IHttpModule
        {
                private ILog logger = LogManager.GetLogger(typeof(PermissionModule));
                //private PassportService passportSrv = null;
                //private NativeServiceManager ServiceManager = null;

                #region IHttpModule 成员

                public void Dispose()
                {

                }

                public void Init(HttpApplication context)
                {
                        //ServiceManager = ServiceManagerFactory.CreateNativeServiceManager();
                        //passportSrv = ServiceManager.CreateService<PassportService>();

                        context.AcquireRequestState += new EventHandler(context_AcquireRequestState);
                        context.Error += new EventHandler(context_Error);
                }

                /// <summary>
                /// 出错处理机制
                /// </summary>
                /// <param name="sender"></param>
                /// <param name="e"></param>
                void context_Error(object sender, EventArgs e)
                {
                        if (ConvertUtil.ToStringOrDefault(ConfigurationManager.AppSettings["global_exception_handler"]).Equals("true"))
                        {
                                HttpContext context = ((HttpApplication)sender).Context;
                                logger.DebugFormat("[YinPSoft-Debug-{0}] - {1}", DateTime.Now.ToString(), context.Error.Message);
                                logger.DebugFormat("[YinPSoft-Debug-{0}] - {1}", DateTime.Now.ToString(), context.Error.StackTrace);

                                context.Response.Redirect("~/" +
                                    ConfigurationManager.AppSettings["ErrorPage"] + "?id=" + context.Error.Message);
                        }
                }

                void context_AcquireRequestState(object sender, EventArgs e)
                {
                        HttpContext context = ((HttpApplication)sender).Context;
                        HttpCookie authCookie = context.Request.Cookies[FormsAuthentication.FormsCookieName];
                        UserInfo userInfo = null;

                        if (authCookie != null)
                        {
                                FormsAuthenticationTicket ticket = FormsAuthentication.Decrypt(authCookie.Value);
                                UserIdentity uid = new UserIdentity(context.User.Identity.Name);
                                if (context.Session != null)
                                {
                                        if (context.Session["USER_INFO"] != null)
                                        {
                                                //uid.UserInfo = context.Session["USER_INFO"] as UserInfo;
                                                userInfo = context.Session["USER_INFO"] as UserInfo;
                                        }
                                        else
                                        {
                                                //如果用户已经通过FORMS验证，但是SESSION中的用户信息不存在
                                                //context.Response.Redirect(FormsAuthentication.LoginUrl+"?status=q&p=__pub__");
                                        }
                                }

                                uid.UserInfo = userInfo;
                                Thread.CurrentPrincipal = new PassportPrincipal(uid);

                                //判断有没有权限访问当前页
                                if (Regex.IsMatch(context.Request.RawUrl,
                                    "/Views/Security\\.*|/Views/Components\\.*|/Views/External\\.*") == false)
                                {
                                        CheckPermission(context);
                                }
                        }
                        else
                        {
                                //authCookie为空则说明未登录或已经注销
                        }
                }
                #endregion

                public void CheckPermission(HttpContext context)
                {
                        UserInfo ui = null;
                        try
                        {
                                if (context.Session != null)
                                {
                                        ui = context.Session["USER_INFO"] as UserInfo;

                                        if (ui != null)
                                        {
                                                string permission = "__pub__";
                                                string querystring = "";

                                                //从参数中取得权限信息，如果没有参数，则将请求中的页面名称作为权限信息
                                                if (context.Request.QueryString.Count > 0)
                                                {
                                                        if (context.Request.QueryString["p"] != null)
                                                        {
                                                                permission = context.Request.QueryString["p"].ToString();        //URL中的权限信息
                                                        }
                                                        //else
                                                        //{
                                                        //    throw new Exception("_missing_permission_argument");
                                                        //}
                                                        querystring = context.Request.Url.Query.Substring(1);//context.Request.RawUrl.Substring(context.Request.RawUrl.IndexOf('?'));
                                                }
                                                else
                                                {
                                                        //throw new Exception("_missing_permission_argument");
                                                        //对权限判断进行严格处理(去掉此处的注释可放宽对权限的判断条件)
                                                        permission = context.Request.Url.Segments[context.Request.Url.Segments.Length - 1].Replace(".aspx", "");
                                                }

                                                if (permission.Equals("__pub__") == false)
                                                {
                                                        DataRow[] rows = ui.Permissions.Select("RESIDENTITY='" + permission + "'");
                                                        if (rows.Length > 0)
                                                        {
                                                                DataRow dr = rows[0];
                                                                context.RewritePath(
                                                                    ConvertUtil.ToStringOrDefault(dr["VIEWNAME"]), "", querystring);
                                                        }
                                                        else
                                                        {
                                                                throw new Exception("_no_permission");
                                                        }
                                                }
                                        }
                                        else
                                        {
                                                //DO NOTHING
                                        }
                                }
                                else
                                {
                                        //DO NOTHING
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }
        }
}
