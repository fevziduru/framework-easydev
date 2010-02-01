using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Principal;

namespace EasyDev.Security
{
    public sealed class Principal : GenericPrincipal
    {
        public Principal(IIdentity identity, string[] roles)
            : base(identity, roles)
        { }

        public Principal(IIdentity identity)
            : base(identity, null)
        { }
    }
}
