using EasyOffice.Models.Word;
using System.Collections.Generic;

namespace EasyOffice.Interfaces
{
    /// <summary>
    /// Word导出Provider
    /// </summary>
    public interface IWordExportProvider
    {
        /// <summary>
        /// 根据模板导出Word文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <returns></returns>
        Word ExportFromTemplate<T>(string templateUrl, T data) where T : class, new();

        /// <summary>
        /// 创建空白Word
        /// </summary>
        /// <returns></returns>
        Word CreateWord();

        /// <summary>
        /// 插入段落
        /// </summary>
        /// <param name="word"></param>
        /// <returns></returns>
        Word InsertParagraphs(Word word, List<Paragraph> paragraphs);

        /// <summary>
        /// 插入表格
        /// </summary>
        /// <param name="word"></param>
        /// <param name="tables"></param>
        /// <returns></returns>
        Word InsertTables(Word word, List<Table> tables);

        /// <summary>
        /// 根据母版表格创建Word
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="datas"></param>
        /// <returns></returns>
        Word CreateFromMasterTable<T>(string templateUrl, IEnumerable<T> datas) where T : class, new();

        /// <summary>
        /// 为文档需求，增加自定义导出Word文档方法，模板对象即包含文章段落占位符替换，也包含表格列表
        /// 针对表格列表：按表格在文档出现顺序，每个表格针对一个列表对象
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <param name="tmpTblIndex">
        /// 标识用于替换数据的表格在模板文档中出现的位置号
        /// 该位置号与模板中(从上往下顺序)表格位置一一对应 从0开始
        /// </param>
        /// <returns></returns>
        Word CreateComplexDoc<T>(string templateUrl, T data, int[] tmpTblIndex) where T : class, new();
    }
}