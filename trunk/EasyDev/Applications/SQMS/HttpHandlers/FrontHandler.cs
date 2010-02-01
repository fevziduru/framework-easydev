using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;

namespace EasyDev.SQMS
{
    public class FrontHandler : IHttpHandler
    {
        private string VirtualUrl = "";

        public FrontHandler(string url)
        {
            VirtualUrl = url;
        }

        #region IHttpHandler 成员

        public bool IsReusable
        {
            get { throw new NotImplementedException(); }
        }

        public void ProcessRequest(HttpContext context)
        {
            context.Server.Transfer(VirtualUrl);            
        }

        #endregion
    }
}
