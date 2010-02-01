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
    public class UserInRole : AbstractModel
	{
		/// <summary>
		/// UserID Field 
		/// </summary>
        [KeyProperty]
		public string UserID { get; set; }

		/// <summary>
		/// RoleID Field 
		/// </summary>
        [KeyProperty]
		public string RoleID { get; set; }
	}
}