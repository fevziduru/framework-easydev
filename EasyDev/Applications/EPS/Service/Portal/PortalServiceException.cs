using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS.Portal
{
    public class PortalServiceException : ExceptionBase,IException
    {
        public PortalServiceException(string msg)
            : base(msg)
        { }
    }
}
