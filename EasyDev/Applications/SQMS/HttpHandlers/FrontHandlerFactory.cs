using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using EasyDev.Util;
using System.Data;
using System.Web.Security;

namespace EasyDev.SQMS
{
    public class FrontHandlerFactory : IHttpHandlerFactory
    {
        #region IHttpHandlerFactory 成员

        public IHttpHandler GetHandler(HttpContext context, string requestType, string url, string pathTranslated)
        {
            return PageParser.GetCompiledPageInstance(url, context.Request.PhysicalPath, context);
        }

        public void ReleaseHandler(IHttpHandler handler)
        {
            
        }

        #endregion
    }
}