using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Security
{
    public class UserSecurityException : ExceptionBase,IException
    {
        public UserSecurityException(string msg)
            : base(msg)
        { }
    }
}
