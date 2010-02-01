using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Util
{
    public class CompressException : ExceptionBase, IException
    {
        public CompressException(string msg)
            : base(msg)
        { }
    }
}
