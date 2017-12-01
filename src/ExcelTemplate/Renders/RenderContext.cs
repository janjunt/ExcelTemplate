using NPOI.SS.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ExcelTemplate.Utility.Extensions;

namespace ExcelTemplate.Renders
{
    /// <summary>
    /// 渲染器上下文
    /// </summary>
    public class RenderContext
    {
        #region 常量
        /// <summary>
        /// 上下文路径分割符
        /// </summary>
        private const string ContextPathSeparator = "|";
        #endregion

        #region 静态字段
        /// <summary>
        /// 属性访问器缓存字典
        /// </summary>
        private static ConcurrentDictionary<Tuple<Type,string>, Func<object, object>> _cachedPropertyAccessor =
            new ConcurrentDictionary<Tuple<Type, string>, Func<object, object>>();
        /// <summary>
        /// 访问器缓存字典
        /// </summary>
        private static ConcurrentDictionary<Tuple<string, string>, Func<RenderContext, object>> _cachedAccessor =
            new ConcurrentDictionary<Tuple<string, string>, Func<RenderContext, object>>();
        #endregion

        #region 属性
        /// <summary>
        /// 上下文路径
        /// </summary>
        public string ContextPath { get; set; }
        /// <summary>
        /// Sheet实例
        /// </summary>
        public ISheet Sheet { get; set; }
        /// <summary>
        /// 父上下文
        /// </summary>
        public RenderContext ParentContext { get; set; }
        /// <summary>
        /// 数据
        /// </summary>
        public object Data { get; set; }
        /// <summary>
        /// 本地变量字典
        /// </summary>
        public IDictionary<string, object> LocalVariables { get; set; } = new Dictionary<string, object>();
        #endregion

        #region 公开方法
        /// <summary>
        /// 清空静态缓存
        /// </summary>
        public static void ClearCache()
        {
            _cachedPropertyAccessor.Clear();
            _cachedAccessor.Clear();
        }

        /// <summary>
        /// 根据名称获取值
        /// </summary>
        /// <param name="name">名称</param>
        /// <returns>名称指定的值</returns>
        public object GetValue(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return Data;
            }

            var accessKey = new Tuple<string, string>(ContextPath, name);
            var accessFunc = _cachedAccessor.GetOrAdd(accessKey, ak =>
            {
                var accessName = ak.Item2;
                return context =>
                {
                    object value;
                    if (context.FirstLocalVariableValue(accessName, out value))
                    {
                        return value;
                    }

                    var dataType = context.Data.GetType();
                    var enumableDefineType = typeof(IEnumerable<>);

                    if (dataType.IsGenericType &&
                        enumableDefineType.MakeGenericType(dataType.GenericTypeArguments[0])
                            .IsAssignableFrom(dataType))
                    {
                        var elementType = dataType.GenericTypeArguments[0];
                        PropertyInfo pi = elementType.GetProperty(accessName);
                        if (pi == null)
                        {
                            if (ParentContext != null)
                            {
                                return context.ParentContext.GetValue(accessName);
                            }

                            throw new ArgumentException($"没有找到指定名称({accessName})的属性");
                        }


                        var sumKey = new Tuple<Type, string>(elementType, accessName);
                        var sumFunc =
                            _cachedPropertyAccessor.GetOrAdd(sumKey, (Tuple<Type, string> pk) =>
                            {
                                var sureEnumableType = typeof(IEnumerable<>).MakeGenericType(elementType);
                                var funcType = Expression.GetFuncType(elementType, pi.PropertyType);
                                var sumMethod = typeof(Enumerable)
                                    .GetGenericMethod(
                                        "Sum",
                                        BindingFlags.Public | BindingFlags.Static,
                                        new[] { elementType },
                                        new[] { sureEnumableType, funcType });

                                var parameterExp = Expression.Parameter(typeof(object), "d");
                                var funcParamter = Expression.Parameter(elementType);
                                var funcLambda = Expression.Lambda(Expression.Property(funcParamter, pi), funcParamter);
                                var dataConvert = Expression.Convert(parameterExp, sureEnumableType);
                                var methodExp = Expression.Call(null, sumMethod, dataConvert, funcLambda);

                                return
                                    Expression.Lambda<Func<object, object>>(
                                        Expression.Convert(methodExp, typeof (object)),
                                        parameterExp)
                                        .Compile();
                            });


                        return sumFunc(context.Data);
                    }
                    else
                    {

                        PropertyInfo pi = dataType.GetProperty(accessName);
                        if (pi == null)
                        {
                            if (context.ParentContext != null)
                            {
                                return context.ParentContext.GetValue(accessName);
                            }

                            throw new ArgumentException($"没有找到指定名称({accessName})的变量或属性");
                        }

                        var propertyKey = new Tuple<Type, string>(dataType, accessName);
                        var propertyFunc =
                            _cachedPropertyAccessor.GetOrAdd(propertyKey, (Tuple<Type, string> pk) =>
                            {
                                var parameterExp = Expression.Parameter(typeof (object), "d");
                                var convertExp = Expression.Convert(parameterExp, dataType);
                                var propertyExp = Expression.Convert(Expression.Property(convertExp, pi),
                                    typeof (object));

                                return Expression.Lambda<Func<object, object>>(propertyExp, parameterExp).Compile();
                            });


                        return propertyFunc(context.Data);
                    }
                };
            });

            return accessFunc(this);
        }

        /// <summary>
        /// 创建子上下文
        /// </summary>
        /// <param name="name">名称</param>
        /// <param name="data">数据</param>
        /// <param name="localVariables">本地变量数组</param>
        /// <returns>子上下文</returns>
        public RenderContext CreateChildContext(
            string name,
            object data,
            params KeyValuePair<string, object>[] localVariables)
        {
            var childContext = new RenderContext()
            {
                ContextPath = GetAccessPath(name),
                Sheet = Sheet,
                Data = data,
                ParentContext = this,
            };
            if (localVariables != null)
            {
                foreach (var kvItem in localVariables)
                {
                    childContext.LocalVariables.Add(kvItem);
                }
            }

            return childContext;
        }

        #endregion

        #region 内部方法

        /// <summary>
        /// 获取本地变量值
        /// </summary>
        /// <param name="name">变量名</param>
        /// <param name="value">输出变量值</param>
        /// <returns>是否找到本地变量</returns>
        protected bool FirstLocalVariableValue(string name, out object value)
        {
            value = null;
            var result = false;
            if (LocalVariables.ContainsKey(name))
            {
                value = LocalVariables[name];
                result = true;
            }
            else if (ParentContext != null)
            {
                result = ParentContext.FirstLocalVariableValue(name, out value);
            }

            return result;
        }

        /// <summary>
        /// 获取访问路径
        /// </summary>
        /// <param name="name">名称</param>
        /// <returns>访问路径</returns>
        private string GetAccessPath(string name)
        {
            return $"{ContextPath}{ContextPathSeparator}{name}";
        }
        #endregion
    }
}
