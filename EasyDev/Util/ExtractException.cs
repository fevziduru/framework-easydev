using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Util
{
    public class ExtractException : ExceptionBase, IException
    {
        public ExtractException(string msg)
            : base(msg)
        { }
    }
}
