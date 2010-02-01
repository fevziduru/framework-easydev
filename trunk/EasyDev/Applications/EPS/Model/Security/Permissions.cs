using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
        [DatabaseName("Oracle")]
        public class Permissions : AbstractModel
        {
                /// <summary>
                /// PermissionID Field 
                /// </summary>
                [KeyProperty]
                public string PERMISSIONID { get; set; }

                /// <summary>
                /// PermissionName Field 
                /// </summary>
                public string PERMISSIONNAME { get; set; }

                /// <summary>
                /// Created Field 
                /// </summary>
                public DateTime CREATED { get; set; }

                /// <summary>
                /// Modified Field 
                /// </summary>
                public DateTime MODIFIED { get; set; }

                /// <summary>
                /// Actived Field 
                /// </summary>
                public string ACTIVED { get; set; }

                /// <summary>
                /// Comments Field 
                /// </summary>
                public string COMMENTS { get; set; }
        }
}