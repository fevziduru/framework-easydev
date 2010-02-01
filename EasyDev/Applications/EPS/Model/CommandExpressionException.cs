using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.EPS
{
    public class CommandExpressionException : ExceptionBase,IException
    {
        public CommandExpressionException(string msg)
            : base(msg)
        { }
    }
}
