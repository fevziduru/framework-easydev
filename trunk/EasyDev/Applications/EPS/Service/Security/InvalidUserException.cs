using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS.Security
{
    public class InvalidUserException : ExceptionBase,IException
    {
        public InvalidUserException(string msg)
            : base(msg)
        { }
    }
}
