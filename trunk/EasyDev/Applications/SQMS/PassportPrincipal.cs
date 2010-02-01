using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Security.Principal;

namespace EasyDev.SQMS
{
    public class PassportPrincipal : GenericPrincipal
    {
        public PassportPrincipal(UserIdentity userIdentity)
            : base(userIdentity, null)
        { }        
    }
}
