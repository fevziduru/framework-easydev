using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

namespace EasyDev.SQMS
{
        [Serializable]
        public class UserInfo : ICloneable
        {
                public string Password { get; set; }

                public string RoleID { get; set; }

                public string RoleName { get; set; }

                public string Passport { get; set; }

                public string EmployeeName { get; set; }

                public DataTable Roles { get; set; }

                public DataTable Permissions { get; set; }

                public string PassportID { get; set; }

                public string EmployeeID { get; set; }

                public string OrganizationID { get; set; }

                public string OrganizationName { get; set; }

                public string DepartmentID { get; set; }

                public string DepartmentName { get; set; }


                #region ICloneable 成员

                /// <summary>
                /// 深度复制用户信息对象
                /// </summary>
                /// <returns></returns>
                public object Clone()
                {
                        UserInfo ui = new UserInfo();
                        ui.DepartmentID = this.DepartmentID;
                        ui.DepartmentName = this.DepartmentName;
                        ui.EmployeeID = this.EmployeeID;
                        ui.EmployeeName = this.EmployeeName;
                        ui.OrganizationID = this.OrganizationID;
                        ui.OrganizationName = this.OrganizationName;
                        ui.Passport = this.Passport;
                        ui.PassportID = this.PassportID;
                        ui.Password = this.Password;
                        ui.Permissions = this.Permissions.Copy();
                        ui.RoleID = this.RoleID;
                        ui.RoleName = this.RoleName;
                        ui.Roles = this.Roles.Copy();

                        return ui;
                }

                #endregion
        }
}