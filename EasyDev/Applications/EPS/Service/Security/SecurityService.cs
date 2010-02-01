using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using EasyDev.EPS;

namespace EasyDev.EPS.Security
{
    public class SecurityService : AbstractService
    {
        public virtual Users FindUser(string userid)
        {
            return GetBO<UserBO>().FindById(userid);
        }

        public virtual Roles FindRole(string roleid)
        {
            return GetBO<RoleBO>().FindById(roleid);
        }

        public virtual Permissions FindPermission(string permissionid)
        {
            return GetBO<PermissionBO>().FindById(permissionid);
        }

        public virtual IList<Users> FindAllUsers()
        {
            return GetBO<UserBO>().FindAll();
        }

        public virtual IList<Roles> FindAllRoles()
        {
            return GetBO<RoleBO>().FindAll();
        }

        public virtual IList<Permissions> FindAllPermissions()
        {
            return GetBO<PermissionBO>().FindAll();
        }

        public virtual bool UpdateUser(Users user)
        {
            return GetBO<UserBO>().Update(user);
        }

        public virtual bool UpdateRole(Roles role)
        {
            return GetBO<RoleBO>().Update(role);
        }

        public virtual bool UpdatePermission(Permissions permission)
        {
            return GetBO<PermissionBO>().Update(permission);
        }

        //public virtual bool DeleteUser(Users user)
        //{
        //    return GetBO<UserBO>().Delete(user);
        //}

        //public virtual bool DeleteRole(Roles role)
        //{
        //    return GetBO<RoleBO>().Delete(role);
        //}

        //public virtual bool DeletePermission(Permissions permission)
        //{
        //    return GetBO<PermissionBO>().Delete(permission);
        //}

        /// <summary>
        /// 创建权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Permissions CreatePermission(Permissions permission)
        {
            try
            {
                if (permission.PERMISSIONID == null
                    || permission.PERMISSIONID.Length == 0)
                {
                    permission.PERMISSIONID = GetDefaultBO<Permissions>().GetNextSequenceId("Permissions");
                }

                permission.CREATED = DateTime.Now;
                permission.MODIFIED = DateTime.Now;
                permission.ACTIVED = "Y";

                if (GetDefaultBO<Permissions>().Save(permission))
                {
                    return permission;
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
        /// 创建权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Permissions CreatePermission(string permissionName)
        {
            Permissions permission = ModelFactory.CreateModel<Permissions>();
            permission.PERMISSIONNAME = permissionName;

            try
            {
                return CreatePermission(permission);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual bool RemovePermission(Permissions permission)
        {
            return GetDefaultBO<Permissions>().Delete(permission);
        }

        /// <summary>
        /// 删除权限
        /// </summary>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual bool RemovePermission(string permissionId)
        {
            Permissions permission = ModelFactory.CreateModel<Permissions>();
            permission.PERMISSIONID = permissionId;

            return RemovePermission(permission);
        }

        /// <summary>
        /// 指定权限的所有角色
        /// </summary>
        /// <param name="permissionId"></param>
        /// <returns></returns>
        public virtual IList<Roles> RolesInPermission(string permissionId)
        {
            return GetBO<PermissionBO>().GetRolesOfPermission(permissionId);
        }

        /// <summary>
        /// 查找权限是否包含指定角色
        /// </summary>
        /// <param name="roleId"></param>
        /// <param name="permission"></param>
        /// <returns></returns>
        public virtual Roles FindRoleInPermission(string roleId, Permissions permission)
        {
            try
            {
                IList<Roles> roleList = GetBO<PermissionBO>().GetRolesOfPermission(permission.PERMISSIONID);

                return (from p in roleList
                        where p.ROLEID == roleId
                        select p).First<Roles>();
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 角色授权
        /// </summary>
        /// <param name="roleId">要授权的角色ID</param>
        /// <param name="permissions">授权列表</param>
        /// <returns>授权失败的权限点，如果返回列表元素个数为0则所有权限授权成功</returns>
        public virtual IList<Permissions> Authorize(string roleId, IList<Permissions> permissions)
        {
            try
            {
                PermissionBO bo = GetBO<PermissionBO>();
                IList<Permissions> failed = new List<Permissions>();

                foreach (Permissions item in permissions)
                {
                    if (bo.GrantRight(roleId, item.PERMISSIONID) == false)
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

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(Users user)
        {
            try
            {
                if (user.USERID == null || user.USERID.Length == 0)
                {
                    user.USERID = GetDefaultBO<Users>().GetNextSequenceId("Users");
                }

                user.CREATED = DateTime.Now;
                user.MODIFIED = DateTime.Now;
                user.ACTIVED = "Y";

                if (GetDefaultBO<Users>().Save(user))
                {
                    return user;
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
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd, string email)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            user.EMAIL = email;
            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 创建用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users CreateUser(string username, string userpwd, string email, string question, string answer)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = userpwd;
            user.EMAIL = email;
            user.QUESTION = question;
            user.ANSWER = answer;

            try
            {
                return CreateUser(user);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool RemoveUser(Users user)
        {
            return GetBO<UserBO>().Delete(user);
        }

        /// <summary>
        /// 删除用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool RemoveUser(string userid)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.USERID = userid;

            return RemoveUser(user);
        }

        /// <summary>
        /// 修改用户密码
        /// </summary>
        /// <param name="user"></param>
        /// <param name="oldpasswd"></param>
        /// <param name="newpasswd"></param>
        public virtual bool ChangeUserPassword(Users user, string newpasswd)
        {
            try
            {
                Users toValidUser = ValidateUser(user);
                if (toValidUser != null)
                {
                    toValidUser.PASSWORD = newpasswd;

                    return GetDefaultBO<Users>().Update(toValidUser);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改用户密码,调用此方法的前提是数据库中的用户表的用户名字段应该有唯一约束
        /// </summary>
        /// <param name="username">用户名（唯一约束）</param>
        /// <param name="oldpasswd"></param>
        /// <param name="newpasswd"></param>
        public virtual bool ChangeUserPassword(string username, string oldpasswd, string newpasswd)
        {
            try
            {
                Users user = ValidateUser(username, oldpasswd);
                if (user != null)
                {
                    user.PASSWORD = newpasswd;
                    return GetDefaultBO<Users>().Update(user);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 修改用户密码
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual bool ChangeUserPassword(string username, string email, string oldpasswd, string newpasswd)
        {
            try
            {
                Users user = ValidateUser(username, oldpasswd, email);
                if (user != null)
                {
                    user.PASSWORD = newpasswd;
                    return GetDefaultBO<Users>().Update(user);
                }
                else
                {
                    throw new InvalidUserException("Invalid_user");
                }
            }
            catch (InvalidUserException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(Users user)
        {
            return GetBO<UserBO>().FindByInstance(user) as Users;
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(string username, string passwd)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.NAME = username;
            user.PASSWORD = passwd;

            return ValidateUser(user);
        }

        /// <summary>
        /// 验证用户
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual Users ValidateUser(string username, string passwd, string email)
        {
            Users user = ModelFactory.CreateModel<Users>();
            user.EMAIL= email;
            user.NAME = username;
            user.PASSWORD = passwd;

            return ValidateUser(user);
        }

        /// <summary>
        /// 用户所在的组
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public virtual IList<Roles> RolesOfUser(string userid)
        {
            return GetBO<RoleBO>().GetRolesByUserID(userid);
        }
    }
}
