using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.GenericList
{
	[DatabaseName("Oracle")]
	public class List : AbstractModel
	{
		/// <summary>
		/// LISTID Field 
		/// </summary>
        [KeyProperty]
		public string LISTID { get; set; }

		/// <summary>
		/// LISTNAME Field 
		/// </summary>
		public string LISTNAME { get; set; }

		/// <summary>
		/// ISSTRONGTYPE Field 
		/// </summary>
		public string ISSTRONGTYPE { get; set; }

		/// <summary>
		/// ISHEADOP Field 
		/// </summary>
		public string ISHEADOP { get; set; }

		/// <summary>
		/// EDITACTION Field 
		/// </summary>
		public string EDITACTION { get; set; }

		/// <summary>
		/// DETAILSACTION Field 
		/// </summary>
		public string DETAILSACTION { get; set; }

		/// <summary>
		/// DELETEACTION Field 
		/// </summary>
		public string DELETEACTION { get; set; }

		/// <summary>
		/// CONTROLLERNAME Field 
		/// </summary>
		public string CONTROLLERNAME { get; set; }

		/// <summary>
        /// MODELTYPE Field 
		/// </summary>
        public string MODELTYPE { get; set; }

		/// <summary>
		/// KEYPROPERTY Field 
		/// </summary>
		public string KEYPROPERTY { get; set; }

		/// <summary>
		/// DATASERVICETYPE Field 
		/// </summary>
		public string DATASERVICETYPE { get; set; }

		/// <summary>
		/// DATAMETHODNAME Field 
		/// </summary>
		public string DATAMETHODNAME { get; set; }

		/// <summary>
		/// CREATED Field 
		/// </summary>
		public DateTime CREATED { get; set; }

		/// <summary>
		/// MODIFIED Field 
		/// </summary>
		public DateTime MODIFIED { get; set; }

		/// <summary>
		/// ACTIVED Field 
		/// </summary>
		public string ACTIVED { get; set; }

        /// <summary>
        /// BOTYPE Field 
        /// </summary>
        public string BOTYPE { get; set; }

        /// <summary>
        /// ISOPVISIBLE Field 
        /// </summary>
        public string ISOPVISIBLE { get; set; }

        /// <summary>
        /// DATASQL Field 
        /// </summary>
        public string DATASQL { get; set; }

        /// <summary>
        /// FROMCLAUSE Field 
        /// </summary>
        public string FROMCLAUSE { get; set; }

        /// <summary>
        /// WHERECLAUSE Field 
        /// </summary>
        public string WHERECLAUSE { get; set; }
	}
}