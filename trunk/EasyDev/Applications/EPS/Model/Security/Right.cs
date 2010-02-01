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
	public class Right : AbstractModel
	{
		/// <summary>
		/// PermissionID Field 
		/// </summary>
        [KeyProperty]
		public string PermissionID { get; set; }

		/// <summary>
		/// RoleID Field 
		/// </summary>
        [KeyProperty]
		public string RoleID { get; set; }

		/// <summary>
		/// Created Field 
		/// </summary>
		public DateTime Created { get; set; }

		/// <summary>
		/// Modified Field 
		/// </summary>
		public DateTime Modified { get; set; }
	}
}