using System;
using System.Collections.Generic;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Resources
{
    public class GlobalizeException : ExceptionBase, IException
    {
        public GlobalizeException(string msg)
            : base(msg)
        {
        }
    }
}
