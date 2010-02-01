using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.Portal
{
	[DatabaseName("Oracle")]
	public class Zone : AbstractModel
	{
		/// <summary>
		/// ZONEID Field 
		/// </summary>
        [KeyProperty]
		public string ZONEID { get; set; }

		/// <summary>
		/// ZONENAME Field 
		/// </summary>
		public string ZONENAME { get; set; }

		/// <summary>
		/// ZONEORDER Field 
		/// </summary>
		public decimal ZONEORDER { get; set; }

		/// <summary>
		/// ZONEDESC Field 
		/// </summary>
		public string ZONEDESC { get; set; }

		/// <summary>
		/// VOID Field 
		/// </summary>
		public string VOID { get; set; }

		/// <summary>
		/// CREATED Field 
		/// </summary>
		public DateTime CREATED { get; set; }

		/// <summary>
		/// MODIFIED Field 
		/// </summary>
		public DateTime MODIFIED { get; set; }

        /// <summary>
        /// ZONESTYLE Field 
        /// </summary>
        public string ZONESTYLE { get; set; }
	}
}