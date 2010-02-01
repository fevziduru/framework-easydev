using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.Cache
{
    [Serializable]
    public class CacheItem
    {
        public string Key { get; set; }

        public object Value { get; set; }

        public CacheItem(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public CacheItem() { }
    }
}
