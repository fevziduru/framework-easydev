using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Portal;
using System.Data;
using EasyDev.PL;
using System.Reflection;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.BusinessObject
{
    /// <summary>
    /// 通用业务对象，此对象中使用的会话对象是默认会话对象
    /// </summary>
    /// <typeparam name="TModel">业务对象对应的模型对象类型</typeparam>
    public class GenericBO<TModel> : AbstractBO where TModel : IModel, new()
    {
        private string _keyField;

        protected override string KeyField
        {
            get
            {
                if (_keyField == null || _keyField.Length == 0)
                {
                    PropertyInfo[] properties = typeof(TModel).GetProperties();
                    PropertyInfo property = null;

                    for (int i = 0; i < properties.Length; i++)
                    {
                        object[] attributes = properties[i].GetCustomAttributes(false);
                        if (attributes.Count(p => p.GetType().Equals(typeof(KeyPropertyAttribute)))>0)
                        {
                            property = properties[i];
                            break;
                        }
                    }

                    if (property != null)
                    {
                        return property.Name;
                    }
                    else
                    {
                        return string.Empty;
                    }
                }
                else
                {
                    return _keyField;
                }
            }
            set
            {
                this._keyField = value;
            }
        }

        protected override void Initialize()
        {
            Model = ModelFactory.CreateModel<TModel>();

            base.Initialize();
        }

        public virtual void SetModel(Type modelType)
        { 
            
        }

        /// <summary>
        /// 保存方法
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public virtual bool Save(TModel entity)
        {
            try
            {
                return DefaultSession.ExecuteCommand(entity.InsertCommand(), entity.ToParamObjects()) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 保存方法（非泛型版本）
        /// </summary>
        /// <param name="entity"></param>
        /// <returns></returns>
        public override bool Save(IModel entity)
        {
            try
            {
                return DefaultSession.ExecuteCommand(entity.InsertCommand(), entity.ToParamObjects()) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public virtual bool Delete(TModel entity)
        {
            try
            {
                return DefaultSession.ExecuteCommand(
                    entity.DeleteCommand(string.Format("{0}=#{0}", KeyField)), entity.ToParamObjects()) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public override bool Delete(IModel entity)
        {
            try
            {
                return DefaultSession.ExecuteCommand(
                    entity.DeleteCommand(string.Format("{0}=#{0}", KeyField)), entity.ToParamObjects()) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public virtual bool Update(TModel entity)
        {
            try
            {
                PropertyInfo result = 
                    entity.GetType().GetProperties().First(p => p.Name.Equals(KeyField, StringComparison.CurrentCultureIgnoreCase));

                string key = string.Empty; ;

                if (result != null)
                {
                    key = result.GetValue(entity, null).ToString();
                }

                IDictionary<string,object> values = entity.ToParamDictionary();

                return DefaultSession.ExecuteCommand(
                    entity.UpdateCommand(string.Format("{0}=#{0}", KeyField.ToLower())), values) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public override bool Update(IModel entity)
        {
            try
            {
                return DefaultSession.ExecuteCommand(
                    entity.UpdateCommand(string.Format("{0}=#{0}", KeyField)), entity.ToParamObjects()) > 0;
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public virtual IList<TModel> FindAll()
        {
            try
            {
                Model = Activator.CreateInstance<TModel>();

                DataTable dt = this.DefaultSession.GetDataTableFromCommand(Model.SelectCommand());
                return ModelFactory.CreateModels<TModel>(dt);
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        public override IList<IModel> FindAllInstance()
        {
            try
            {
                EnsureConcreteBO();

                if (Model == null)
                {
                    throw new ModelNullException("Model_is_not_initialized");
                }

                DataTable dt = this.DefaultSession.GetDataTableFromCommand(Model.SelectCommand());

                return ModelFactory.CreateModels(dt, typeof(TModel));
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        private void EnsureConcreteBO()
        {
            if (this.GetType().Equals(typeof(GenericBO<TModel>)))
            {
                throw new NotSupportedException(@"This method is not supported in GenericBO. 
                    Please use a concrete BO which inherit from GenericBO, 
                    such as UserBO, RoleBO, NewsBO and so forth.");
            }
        }

        /// <summary>
        /// 查找实例是否存在
        /// </summary>
        /// <typeparam name="TInstance"></typeparam>
        /// <param name="instance"></param>
        /// <returns></returns>
        public virtual TInstance FindByInstance<TInstance>(TInstance instance) 
            where TInstance : AbstractModel, new()
        {
            try
            {
                string command = SQLBuilder.BuildSelectInstance(instance);
                DataTable dt = DefaultSession.GetDataTableFromCommand(command);

                if (dt.Rows.Count > 0)
                {
                    return ModelFactory.CreateModel<TInstance>(dt);
                }
                else
                {
                    return null;
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
            catch (ModelNotFoundException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 查找实例是否存在
        /// </summary>
        /// <typeparam name="TInstance"></typeparam>
        /// <param name="instance"></param>
        /// <returns></returns>
        public override IModel FindByInstance(IModel instance)
        {
            try
            {
                string command = SQLBuilder.BuildSelectInstance(instance);
                DataTable dt = DefaultSession.GetDataTableFromCommand(command);
                if (dt.Rows.Count > 0)
                {
                    return ModelFactory.CreateModel(dt, Model.GetType());
                }
                else
                {
                    return null;
                }
            }
            catch (DALCoreException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 根据KEYFIELD查找模型对象
        /// </summary>
        /// <typeparam name="TModel"></typeparam>
        /// <param name="id"></param>
        /// <returns></returns>
        public virtual TModel FindById(string id)
        {
            try
            {
                EnsureConcreteBO();

                string cmd = Model.SelectCommand(string.Format("{0}=#{0}", KeyField));
                DataTable dt = DefaultSession.GetDataTableFromCommand(cmd, id);

                return ModelFactory.CreateModel<TModel>(dt);
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}
