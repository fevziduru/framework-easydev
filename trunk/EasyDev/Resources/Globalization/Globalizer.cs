using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace EasyDev.Resources
{
    public class Globalizer : GlobalizerBase
    {
        public Globalizer(XmlDocument languageItems)
            : base(languageItems)
        {
        }

        public override string GetString(string key)
        {
            return this.GetValueByKey(key);
        }

        public override string GetFormatString(string key, params object[] objs)
        {
            return string.Format(this.GetValueByKey(key), objs);
        }
    }
}
