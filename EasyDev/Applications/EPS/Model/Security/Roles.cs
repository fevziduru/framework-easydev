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
	public class Roles : AbstractModel
	{
		/// <summary>
		/// RoleID Field 
		/// </summary>
        [KeyProperty]
        [RequiredValidator]
		public string ROLEID { get; set; }

		/// <summary>
		/// RoleName Field 
		/// </summary>
        [RequiredValidator]
		public string ROLENAME { get; set; }

		/// <summary>
		/// Comments Field 
		/// </summary>
		public string COMMENTS { get; set; }

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
	}
}