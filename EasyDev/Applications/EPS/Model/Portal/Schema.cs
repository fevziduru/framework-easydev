using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.Portal
{
	[DatabaseName("Oracle")]
	public class Schema : AbstractModel
	{
		/// <summary>
		/// SCHEMAID Field 
		/// </summary>
        [KeyProperty]
		public string SCHEMAID { get; set; }

		/// <summary>
		/// SCHEMANAME Field 
		/// </summary>
		public string SCHEMANAME { get; set; }

		/// <summary>
		/// COMMENTS Field 
		/// </summary>
		public string COMMENTS { get; set; }

        /// <summary>
        /// ISDEFAULT Field 
        /// </summary>
        public string ISDEFAULT { get; set; }
	}
}