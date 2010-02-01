using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Security.Principal;

namespace EasyDev.SQMS
{
        public class UserIdentity : GenericIdentity
        {
                public UserIdentity(string name)
                        : base(name)
                { }

                //public string Password { get; set; }
                public UserInfo UserInfo { get; set; }
        }
}
