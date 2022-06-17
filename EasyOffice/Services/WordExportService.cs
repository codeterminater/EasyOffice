using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EasyOffice.Services
{
    public class WordExportService : IWordExportService
    {
        private readonly IWordExportProvider _wordExportProvider;

        public WordExportService(IWordExportProvider wordExportProvider)
        {
            _wordExportProvider = wordExportProvider;
        }

        public Task<Word> CreateWordAsync(IEnumerable<IWordElement> elements, IWordExportProvider customWordExportProvider = null)
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;

            var word = provider.CreateWord();

            elements?.ToList().ForEach(x =>
            {
                if (x is Paragraph)
                {
                    word = provider.InsertParagraphs(word, new List<Paragraph>() { x as Paragraph });
                }

                if (x is Table)
                {
                    word = provider.InsertTables(word, new List<Table>() { x as Table });
                }
            });

            return Task.FromResult(word);
        }

        public Task<Word> CreateFromTemplateAsync<T>(string templateUrl
            , T wordData
            , IWordExportProvider customWordExportProvider = null)
        where T : class, new()
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;

            var word = _wordExportProvider.ExportFromTemplate(templateUrl, wordData);
            return Task.FromResult(word);
        }

        public Task<Word> CreateFromMasterTableAsync<T>(string templateUrl
            , IEnumerable<T> datas
            , IWordExportProvider customWordExportProvider = null)
            where T : class, new()
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;
            var word = _wordExportProvider.CreateFromMasterTable(templateUrl, datas);
            return Task.FromResult(word);
        }

        /// <summary>
        /// 根据模板创建复杂文档：含有段落文字，或同时含有表格数据的文档
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="templateUrl"></param>
        /// <param name="data"></param>
        /// <param name="tableIdx">模板文档中，要替换数据的模板表格在文档中从上至下出现的顺序号</param>
        /// <param name="customWordExportProvider"></param>
        /// <returns></returns>

        public Task<Word> CreateComplexWordAsync<T>(string templateUrl,
            T wordData,
            int[] tableIdx,
            IWordExportProvider customWordExportProvider = null)
            where T : class, new()
        {
            var provider = customWordExportProvider == null ? _wordExportProvider : customWordExportProvider;

            var word = _wordExportProvider.CreateComplexDoc(templateUrl, wordData, tableIdx);
            return Task.FromResult(word);
        }
    }
}