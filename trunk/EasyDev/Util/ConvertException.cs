using System;
using System.Collections.Generic;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Util
{
    public class ConvertException : ExceptionBase, IException
    {
        public ConvertException(string msg)
            : base(msg)
        { }
    }
}
