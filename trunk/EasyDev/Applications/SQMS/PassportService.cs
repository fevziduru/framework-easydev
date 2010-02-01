using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.BL;
using System.Data;
using EasyDev.Util;
using EasyDev.SQMS;

namespace EasyDev.SQMS
{
        public class PassportService : GenericService
        {
                //private RoleService roleService = null;        

                protected override void Initialize()
                {
                        this.BOName = "PASSPORT";
                        //roleService = ServiceManager.CreateService<RoleService>();

                        base.Initialize();
                }

                /// <summary>
                /// 取得所有角色
                /// </summary>
                /// <returns></returns>
                //public DataSet GetRoles()
                //{
                //    return roleService.LoadAll();
                //}

                /// <summary>
                /// 用户登录验证
                /// </summary>
                /// <param name="name">用户名</param>
                /// <param name="password">密码</param>
                /// <returns></returns>
                public UserInfo Login(string name, string password, string role)
                {
                        try
                        {
                                return GetUserInfo(name, password, role);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }                

                public UserInfo GetUserInfo(string name, string password, string role)
                {
                        UserInfo ui = null;
                        DataRow userRow = null;
                        try
                        {
                                if (password != null)
                                {
                                        userRow = DataSetUtil.GetFirstRowFromDataSet(LoadByCondition(
                                            string.Format("PASSPORT='{0}' AND PASSWORD='{1}'", name, password)), this.BOName);
                                }
                                else
                                {
                                        userRow = DataSetUtil.GetFirstRowFromDataSet(LoadByCondition(
                                            string.Format("PASSPORT='{0}'", name)), this.BOName);
                                }

                                if (userRow != null)
                                {
                                        //取账号信息
                                        //DataRow drPassport =
                                        //    DataSetUtil.GetFirstRowFromDataSet(
                                        //    LoadByCondition("ISVOID='N' AND PASSPORT='" + name + "'"), BOName);

                                        //取职员信息
                                        DataRow drUserInfo = DataSetUtil.GetFirstRowFromDataTable(
                                            DefaultSession.GetDataTableFromCommand(
                                                @"select  t.empid, 
                                        t.empname, 
                                        d.orgid deptid,
                                        d.orgname deptname,
                                        o.orgid orgid, 
                                        o.orgname orgname, 
                                        o.orgalias,
                                        r.roleid,
                                        r.rolename,
                                        p.passportid,
                                        p.passport
                                from employee t
                                left join ORGANIZATION d on t.deptid = d.orgid and d.orgtype='department'
                                left join ORGANIZATION o on o.orgid = t.organizationid and o.orgtype='organization'
                                left join passport p on p.empid = t.empid and p.isvoid = 'N'
                                left join userrole ur on p.passportid = ur.passportid
                                left join role r on r.roleid = ur.roleid
                                where p.passport = :passport and r.roleid=:roleid", name, role));
                                        if (drUserInfo != null)
                                        {
                                                ui = new UserInfo();
                                                //ui.Passport = name;
                                                //ui.RoleID = role;
                                                ui.Passport = ConvertUtil.ToStringOrDefault(drUserInfo["PASSPORT"]);
                                                ui.PassportID = ConvertUtil.ToStringOrDefault(drUserInfo["PASSPORTID"]);
                                                ui.EmployeeID = ConvertUtil.ToStringOrDefault(drUserInfo["EMPID"]);
                                                ui.OrganizationID = ConvertUtil.ToStringOrDefault(drUserInfo["orgid"]);
                                                ui.OrganizationName = ConvertUtil.ToStringOrDefault(drUserInfo["orgname"]);
                                                ui.DepartmentID = ConvertUtil.ToStringOrDefault(drUserInfo["deptid"]);
                                                ui.DepartmentName = ConvertUtil.ToStringOrDefault(drUserInfo["deptname"]);
                                                ui.EmployeeName = ConvertUtil.ToStringOrDefault(drUserInfo["EMPNAME"]);
                                                ui.RoleID = ConvertUtil.ToStringOrDefault(drUserInfo["ROLEID"]);
                                                ui.RoleName = ConvertUtil.ToStringOrDefault(drUserInfo["ROLENAME"]);

                                                ui.Roles = GetUserRole(name);   //取得用户角色
                                                ui.Permissions = GetUserPermission(name, role); //取得用户权限以相关的操作
                                        }
                                        else
                                        {
                                                return null;
                                        }
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }

                        return ui;
                }

                /// <summary>
                /// 取得用户的角色
                /// </summary>
                /// <param name="passport"></param>
                /// <returns></returns>
                public DataTable GetUserRole(string passport)
                {
                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(
                                        @"select r.rolename, r.roleid from passport p
                        left join userrole ur on p.passportid = ur.passportid
                        left join role r on r.roleid = ur.roleid
                        where p.passport = :passport and p.isvoid = 'N'", passport);
                                dt.TableName = "UserRole";

                                return dt;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 取得用户权限
                /// </summary>
                /// <param name="passport"></param>
                /// <returns></returns>
                public DataTable GetUserPermission(string passport, string role)
                {
                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(
                                        @"select ri.resname,ri.restype,o.opidentity, ri.viewname, ri.resid, ri.residentity, o.opid, o.opname, ri.location
                        from passport p
                        left join userrole ur on p.passportid = ur.passportid
                        left join role r on r.roleid = ur.roleid
                        left join respermission rp on rp.roleid = r.roleid
                        left join resourceitem ri on ri.resid = rp.resid
                        left join operation o on o.opid = rp.opid and o.isvoid='N' and o.oporder >= 0
                        where p.passport = :passport and p.isvoid = 'N' and r.roleid=:roleid 
                        order by o.oporder asc", passport, role);
                                dt.TableName = "UserPermission";

                                return dt;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                public string GetPassport(string empid)
                {
                        try
                        {
                                return ConvertUtil.ToStringOrDefault(DefaultSession.GetScalarObjectFromCommand(
                                    @"select passport from passport where empid=:empid and isvoid='N'", empid));
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                public DataSet GetPassportByEmployeeID(string empid)
                {
                        try
                        {
                                DataSet ds = DefaultSession.GetDataSetFromCommand("select * from passport where passport.empid=:empid", empid);
                                ds.Tables[0].TableName = "PASSPORT";
                                return ds;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                public bool IsPassportExist(string passport)
                {
                        try
                        {
                                DataSet ds = LoadByCondition("passport='" + passport + "'");
                                return (ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }
        }
}
