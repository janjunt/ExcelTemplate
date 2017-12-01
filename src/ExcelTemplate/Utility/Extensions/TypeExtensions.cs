using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace ExcelTemplate.Utility.Extensions
{
    /// <summary>
    /// Type扩展
    /// </summary>
    public static class TypeExtensions
    {
        /// <summary>
        /// 获取泛型方法
        /// </summary>
        /// <param name="type">Type实例</param>
        /// <param name="name">方法名称</param>
        /// <param name="bindingAttr">绑定标志及成员搜索方式</param>
        /// <param name="typeArguments">泛型类型</param>
        /// <param name="parameterTypeArguments">参数类型</param>
        /// <returns></returns>
        public static MethodInfo GetGenericMethod(this Type type, 
            string name, 
            BindingFlags bindingAttr,
            IEnumerable<Type> typeArguments, 
            IEnumerable<Type> parameterTypeArguments)
        {
            var methods = type.GetMethods(bindingAttr).Where(m => m.Name == name && m.IsGenericMethod);
            MethodInfo findMethod = null;
            var parameterTypes = parameterTypeArguments.ToList();
            foreach (var method in methods)
            {
                var sureMethod = method.MakeGenericMethod(typeArguments.ToArray());
                var parameters = sureMethod.GetParameters();
                if (parameters.Length != parameterTypes.Count)
                {
                    continue;
                }

                var isFind = true;
                for (var i = 0; i < parameters.Length; i++)
                {
                    if (parameters[i].ParameterType != parameterTypes[i])
                    {
                        isFind = false;
                        break;
                    }
                }

                if (isFind)
                {
                    findMethod = sureMethod;
                    break;
                }
            }

            return findMethod;
        }
    }
}
