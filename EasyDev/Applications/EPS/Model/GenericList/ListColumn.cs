using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.GenericList 
{
	[DatabaseName("Oracle")]
	public class ListColumn : AbstractModel
	{
		/// <summary>
		/// COLUMNID Field 
		/// </summary>
		public string COLUMNID { get; set; }

		/// <summary>
		/// COLUMNNAME Field 
		/// </summary>
		public string COLUMNNAME { get; set; }

		/// <summary>
		/// DISPLAYNAME Field 
		/// </summary>
		public string DISPLAYNAME { get; set; }

		/// <summary>
		/// DATATYPE Field 
		/// </summary>
		public string DATATYPE { get; set; }

		/// <summary>
		/// LISTID Field 
		/// </summary>
		public string LISTID { get; set; }

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
        /// ALIAS Field 
        /// </summary>
        public string ALIAS { get; set; }

        /// <summary>
        /// ISGROUP Field 
        /// </summary>
        public string ISGROUP { get; set; }

        /// <summary>
        /// ORDERTYPE Field 
        /// </summary>
        public string ORDERTYPE { get; set; }

        public override string ToString()
        {
            StringBuilder col = new StringBuilder();
            if (this.COLUMNNAME != null && this.COLUMNNAME.Length > 0)
            {
                if (this.ALIAS != null && this.ALIAS.Length > 0)
                {
                    col.AppendFormat(" {0} {1} ",COLUMNNAME, ALIAS);
                }
                else
                {
                    col.AppendFormat(COLUMNNAME);
                }

                //if (ORDERTYPE != null && 
                //    string.Equals(this.ORDERTYPE, "NONE",StringComparison.CurrentCultureIgnoreCase) == false)
                //{
                //    col.Append(ORDERTYPE);
                //}
            }
            else
            {
                throw new Exception("column name is null");
            }

            return col.ToString();
        }
	}
}