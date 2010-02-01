using System;
using System.Collections.Generic;
using System.Text;
using EasyDev.Common;

namespace EasyDev.Configuration
{
    public class ConfigException : ExceptionBase, IException
    {
        public ConfigException(string msg)
            : base(msg)
        {
        }
    }
}
