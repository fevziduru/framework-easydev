using System;
using System.Collections.Generic;
using System.Text;

namespace EasyDev.Resources
{
    public interface IGlobalizer
    {
        string GetString(string key);
        
        string GetFormatString(string key, params object[] objs);
    }
}
