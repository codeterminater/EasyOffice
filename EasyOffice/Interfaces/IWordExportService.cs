using EasyOffice.Models.Word;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Word导出服务
    /// </summary>
    public interface IWordExportService
    {
        /// <summary>
        /// 根据模板创建Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Task<Word> CreateFromTemplateAsync<T>(string templateUrl, T data, IWordExportProvider customWordExportProvider = null) where T : class, new();

        /// <summary>
        /// 从空白创建Word文档
        /// </summary>
        /// <param name="elements"></param>
        /// <returns></returns>
        Task<Word> CreateWordAsync(IEnumerable<IWordElement> elements, IWordExportProvider customWordExportProvider = null);

        /// <summary>
        /// 从母版表格中复制样式，N条数据生成N个复制品并替换数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        Task<Word> CreateFromMasterTableAsync<T>(string templateUrl, IEnumerable<T> datas, IWordExportProvider customWordExportProvider = null)
            where T : class, new();

        /// <summary>
        /// 根据模板创建复杂文档：含有段落文字，或同时含有表格数据的文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <param name="tableIdx">模板文档中，要替换数据的模板表格在文档中从上至下出现的顺序号</param>
        /// <param name="customWordExportProvider"></param>
        /// <returns></returns>
        Task<Word> CreateComplexWordAsync<T>(string templateUrl, T data, int[] tableIdx, IWordExportProvider customWordExportProvider = null)
            where T : class, new();
    }
}