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
	public class Users : AbstractModel
	{
		/// <summary>
		/// UserID Field 
		/// </summary>
        [KeyProperty]
        [RequiredValidator(ErrorMessage = "用户ID为必填项")]
		public string USERID { get; set; }

		/// <summary>
		/// Name Field 
		/// </summary>
        [RequiredValidator(ErrorMessage = "用户名为必填项")]
		public string NAME { get; set; }

		/// <summary>
		/// Created Field 
		/// </summary>
		public DateTime CREATED { get; set; }

		/// <summary>
		/// Code Field 
		/// </summary>
		public string CODE { get; set; }

		/// <summary>
		/// Modified Field 
		/// </summary>
		public DateTime MODIFIED { get; set; }

		/// <summary>
		/// Email Field 
		/// </summary>
        [EmailValidator(ErrorMessage = "EMAIL格式不对")]
		public string EMAIL { get; set; }

		/// <summary>
		/// Password Field 
		/// </summary>
		public string PASSWORD { get; set; }

		/// <summary>
		/// Question Field 
		/// </summary>
		public string QUESTION { get; set; }

		/// <summary>
		/// Answer Field 
		/// </summary>
		public string ANSWER { get; set; }

		/// <summary>
		/// Comments Field 
		/// </summary>
		public string COMMENTS { get; set; }

		/// <summary>
		/// Actived Field 
		/// </summary>
		public string ACTIVED { get; set; }
	}
}