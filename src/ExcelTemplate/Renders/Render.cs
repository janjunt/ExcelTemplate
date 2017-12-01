using System.Collections.Generic;

namespace ExcelTemplate.Renders
{
    /// <summary>
    /// 渲染器接口
    /// </summary>
    public interface IRender
    {
        /// <summary>
        /// 名称
        /// </summary>
        string Name { get; set; }
        /// <summary>
        /// 子渲染器列表
        /// </summary>
        IList<IRender> Childs { get; set; }
        /// <summary>
        /// 渲染
        /// </summary>
        /// <param name="context">渲染下下文</param>
        void Render(RenderContext context);
    }
}
