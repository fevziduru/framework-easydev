using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS.BusinessObject;
using EasyDev.EPS.GenericList;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
    public class RoleService : AbstractService, IListDataService
    {
        protected override void Initialize()
        {
            base.Initialize();
        }

        /// <summary>
        /// 创建角色
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual Roles CreateRole(Roles role)
        {
            try
            {
                if (role.ROLEID == null || role.ROLENAME.Length == 0)
                {
                    role.ROLEID = GetDefaultBO<Roles>().GetNextSequenceId("Roles");
                }

                role.CREATED = DateTime.Now;
                role.MODIFIED = DateTime.Now;
                role.ACTIVED = "Y";

                if (GetDefaultBO<Roles>().Save(role))
                {
                    return role;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception e)
            {
                throw e;
            }  
        }

        /// <summary>
        /// 创建角色
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual Roles CreateRole(string rolename)
        {
            Roles role = ModelFactory.CreateModel<Roles>();
            role.ROLENAME = rolename;

            try
            {
                return CreateRole(role);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除角色
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual bool RemoveRole(Roles role)
        {
            return GetDefaultBO<Roles>().Delete(role);
        }

        /// <summary>
        /// 删除角色
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual bool RemoveRole(string roleId)
        {
            Roles role = ModelFactory.CreateModel<Roles>();
            role.ROLEID = roleId;

            return RemoveRole(role);
        }

        /// <summary>
        /// 取得角色中的所有用户
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual IList<Users> FindAllUsersInRole(Roles role)
        {
            return GetBO<UserBO>().GetUsersByRoleID(role.ROLEID);
        }

        /// <summary>
        /// 在指定组中查找特定用户
        /// </summary>
        /// <param name="user"></param>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual Users FindUserInRole(string userId, Roles role)
        {
            try
            {
                //找到指定角色下的所有用户
                IList<Users> userList = GetBO<UserBO>().GetUsersByRoleID(role.ROLEID);

                return (from p in userList
                        where p.USERID == userId
                        select p).First<Users>();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 查找角色所拥有的权限
        /// </summary>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual IList<Permissions> FindPermissionsOfRole(string roleId)
        {
            return GetBO<RoleBO>().GetPermissionsOfRole(roleId);
        }

        /// <summary>
        /// 查找角色是否有指定的权限
        /// </summary>
        /// <param name="permission"></param>
        /// <param name="role"></param>
        /// <returns></returns>
        public virtual Permissions FindPermissionInRole(string permissionId, Roles role)
        {
            try
            {
                IList<Permissions> permissionList = GetBO<RoleBO>().GetPermissionsOfRole(role.ROLEID);

                return (from p in permissionList 
                        where p.PERMISSIONID == permissionId 
                        select p).First<Permissions>();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 分配角色
        /// </summary>
        /// <param name="userid"></param>
        /// <param name="roles"></param>
        /// <returns>分配失败的角色，如果返回列表元素个数为0则所有角色分配成功</returns>
        public virtual IList<Roles> AssignRoles(string userid, IList<Roles> roles)
        {
            try
            {
                IList<Roles> failed = new List<Roles>();
                RoleBO bo = GetBO<RoleBO>();

                foreach (Roles item in roles)
                {
                    if (bo.Assign(userid, item.ROLEID) == false)
                    {
                        failed.Add(item);
                    }
                }

                return failed;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        #region IListDataService 成员

        public object DefaultDataMethod()
        {
            return GetBO<RoleBO>().GetAllRoles();
        }

        #endregion
    }
}
