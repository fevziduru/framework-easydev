using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using System.Data;
using EasyDev.EPS.Portal;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
        public class RoleBO : GenericBO<Roles>
        {
                /// <summary>
                /// 创建业务对象的时候，首先要覆盖基类的Initialize方法实现模型和KeyField属性的初始化
                /// </summary>
                protected override void Initialize()
                {
                        Model = ModelFactory.CreateModel<Roles>();
                        KeyField = "ROLEID";

                        base.Initialize();
                }

                public DataTable GetAllRoles()
                {
                        return DefaultSession.GetDataTableFromCommand("SELECT ROLEID ,ROLENAME FROM Roles");
                }

                /// <summary>
                /// 根据用户ID取得指定用户所属的所有角色
                /// </summary>
                /// <param name="userId"></param>
                /// <returns></returns>
                public virtual IList<Roles> GetRolesByUserID(string userId)
                {
                        //TODO:SQL命令通过对象化方式构建
                        string command = string.Format(
                            @"SELECT ROLES.* 
                  FROM USERINROLE 
                  LEFT JOIN ROLES ON USERINROLE.ROLEID=ROLES.ROLEID WHERE USERINROLE.USERID='{0}'", userId);

                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(command);
                                return ModelFactory.CreateModels<Roles>(dt);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 取得角色所拥有的权限
                /// </summary>
                /// <param name="roleId"></param>
                /// <returns></returns>
                public virtual IList<Permissions> GetPermissionsOfRole(string roleId)
                {
                        string command = string.Format(@"SELECT PERMISSIONS.* 
                FROM RIGHT 
                LEFT JOIN PERMISSIONS ON PERMISSIONS.PERMISSIONID=RIGHT.PERMISSIONID 
                WHERE RIGHT.ROLEID='{0}'",
                            roleId);

                        try
                        {
                                DataTable dt = DefaultSession.GetDataTableFromCommand(command);

                                return ModelFactory.CreateModels<Permissions>(dt);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 保存角色分配信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool SaveAssignment(string userid, string roleid)
                {
                        string command = string.Format(@"INSERT INTO USERINROLE (USERID,ROLEID,ACTIVED) 
                VALUES('{0}','{1}','Y')", userid, roleid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) > 0;
                        }
                        catch (Exception e)
                        {
                                throw;
                        }
                }

                /// <summary>
                /// 删除角色分配信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool DeleteAssignment(string userid, string roleid)
                {
                        string command = string.Format(@"DELETE FROM USERINROLE WHERE USERID='{0}' AND ROLEID='{1}'",
                            userid, roleid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) >= 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 更新角色分配信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool UpdateAssignment(string userid, string roleid)
                {
                        string command = string.Format(@"UPDATE USERINROLE
                SET USERID='{0}', ROLEID='{1}' 
                WHERE USERID='{0}' AND ROLEID='{1}'", userid, roleid);

                        try
                        {
                                return DefaultSession.ExecuteCommand(command) >= 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 查找角色分配信息
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionId"></param>
                /// <returns></returns>
                public virtual bool FindAssignment(string userid, string roleid)
                {
                        string command = string.Format(@"SELECT Count(*) 
                FROM USERINROLE WHERE USERID='{0}' AND ROLEID='{1}'",
                            userid, roleid);

                        try
                        {
                                return Convert.ToInt32(DefaultSession.GetScalarObjectFromCommand(command)) > 0;
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 角色分配
                /// </summary>
                /// <param name="roleid"></param>
                /// <param name="permissionid"></param>
                /// <returns></returns>
                public virtual bool Assign(string userid, string roleid)
                {
                        try
                        {
                                if (FindAssignment(userid, roleid))
                                {
                                        return UpdateAssignment(userid, roleid);
                                }
                                else
                                {
                                        return SaveAssignment(userid, roleid);
                                }
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }
                }
        }
}
