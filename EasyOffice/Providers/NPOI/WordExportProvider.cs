using EasyOffice.Attributes;
using EasyOffice.Enums;
using EasyOffice.Interfaces;
using EasyOffice.Models.Word;
using EasyOffice.Utils;
using NPOI.XWPF.UserModel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace EasyOffice.Providers.NPOI
{
    public class WordExportProvider : IWordExportProvider
    {
        #region Field

        private static uint picId = 0;
        private const string ParagraphKeyName = "Graphs";

        #endregion Field

        #region Utils

        /// <summary>
        /// 图片Id
        /// </summary>
        private static uint PicId
        {
            get
            {
                picId++;
                return picId;
            }
        }

        /// <summary>
        /// 替换占位符
        /// </summary>
        /// <param name="word"></param>
        private void ReplacePlaceholders<T>(XWPFDocument word, T wordData)
            where T : class, new()
        {
            if (word == null)
            {
                throw new ArgumentNullException("word");
            }

            Dictionary<string, string> stringReplacements = WordHelper.GetReplacements(wordData);

            Dictionary<string, IEnumerable<Picture>> pictureReplacements = WordHelper.GetPictureReplacements(wordData);

            NPOIHelper.ReplacePlaceholdersInWord(word, stringReplacements, pictureReplacements);
        }

        /// <summary>
        /// 获取 占位符：值 字典
        /// 对包含即有数据列表(对文档模板表格使用)，也有普通数据(对文章段落)
        /// 只替换模型对象中字符串类型占位符, 不处理图片类型，因为自定义业务中不包含插入图片需求
        /// for model class like:
        ///     class Model {
        ///         string foo;
        ///         int bar;
        ///         double val1;
        ///         IList<BaseTableEle> list;(/ICollection/IEnumerable/IEnumerable/ICollection/IReadOnlyList/IReadOnlyCollection)
        ///     }
        ///
        /// For Results like:
        ///  graph[0] :  Key0: value0
        ///              Key1: value1 ...
        ///  PQITableDatas:
        ///               [0] : Key0: value0
        ///                     Key1: value1 ...
        ///
        ///               [1] : Key0: value0
        ///                     Key1: value1 ...
        ///  MQITableDatas: same as PQITableDatas
        ///
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, List<Dictionary<string, string>>> GeneratePlacehoderDict4ComplexType<T>(T wordData)
           where T : class, new()
        {
            Dictionary<string, List<Dictionary<string, string>>> replacements = new Dictionary<string, List<Dictionary<string, string>>>
            {
                { ParagraphKeyName, new List<Dictionary<string, string>>()
                    {
                        new Dictionary<string, string>()
                    }
                }
            };
            Type type = typeof(T);
            PropertyInfo[] props = type.GetProperties();
            foreach (PropertyInfo prop in props)
            {
                if (prop.PropertyType == typeof(Picture) || typeof(IEnumerable<Picture>).IsAssignableFrom(prop.PropertyType))
                    continue;
                if (prop.PropertyType == typeof(string) || prop.PropertyType == typeof(int) || prop.PropertyType == typeof(double) || prop.PropertyType == typeof(long) || prop.PropertyType == typeof(decimal))
                {
                    //所有非LIst(即普通类型数据：主要用于文章段落占位符)放入Graphs下
                    if (replacements.TryGetValue(ParagraphKeyName, out var graphList))
                    {
                        var replacement = prop.GetValue(wordData)?.ToString();
                        var placeholder = prop.IsDefined(typeof(PlaceholderAttribute))
                                        ? prop.GetCustomAttribute<PlaceholderAttribute>().Placeholder.ToString()
                                        : "{" + prop.Name + "}";
                        graphList[0].Add(placeholder, replacement);
                    }
                }
                else
                {
                    var enumerable = prop.GetValue(wordData) as IEnumerable;
                    if (enumerable == null)
                        continue;

                    Type[] interfaces = prop.PropertyType.GetInterfaces();//enumerable.GetType().GetInterfaces();
                    Type elementType = (from i in interfaces
                                        where i.IsGenericType && i.GetGenericTypeDefinition() == typeof(IEnumerable<>)
                                        select i.GetGenericArguments()[0]).FirstOrDefault();

                    if (elementType.IsSubclassOf(typeof(BaseTableEle)))
                    {
                        string enumerablePrefix = prop.Name;
                        List<Dictionary<string, string>> listDict = new List<Dictionary<string, string>>();
                        foreach (var tblEle in enumerable)
                        {
                            var tableEle = tblEle as BaseTableEle;
                            var keyValueList = tableEle.GetKeyValuePairs();
                            if (keyValueList.Count > 0)
                            {
                                Dictionary<string, string> dict = new Dictionary<string, string>();
                                foreach (var entry in keyValueList)
                                {
                                    string newKey = "{" + $"{enumerablePrefix}.{entry.Key}" + "}";
                                    dict.Add(newKey, entry.Value);
                                }
                                listDict.Add(dict);
                            }
                        }
                        replacements.Add(enumerablePrefix, listDict);
                    }
                }
            }

            return replacements;
        }

        private XWPFDocument InsertTable(XWPFDocument doc, Table t)
        {
            var maxColCount = t.Rows.Max(x => x.Cells.Count);

            if (t == null) return doc;

            var table = doc.CreateTable();

            table.Width = t.Width;

            int index = 0;
            t.Rows?.ForEach(r =>
            {
                XWPFTableRow tableRow = index == 0 ? table.GetRow(0) : table.CreateRow();

                for (int i = 0; i < r.Cells.Count; i++)
                {
                    var cell = r.Cells[i];
                    var xwpfCell = i == 0 ? tableRow.GetCell(0) : tableRow.AddNewTableCell();
                    foreach (var para in cell.Paragraphs)
                    {
                        xwpfCell.AddParagraph().Set(para);
                    }

                    if (!string.IsNullOrWhiteSpace(cell.Color))
                    {
                        tableRow.GetCell(i).SetColor(cell.Color);
                    }
                }

                //补全单元格，并合并
                var rowColsCount = tableRow.GetTableICells().Count;
                if (rowColsCount < maxColCount)
                {
                    for (int i = rowColsCount - 1; i < maxColCount; i++)
                    {
                        tableRow.CreateCell();
                    }
                    tableRow.MergeCells(rowColsCount - 1, maxColCount);
                }

                index++;
            });

            return doc;
        }

        private XWPFDocument InsertParagraph(XWPFDocument doc, Paragraph paragraph)
        {
            doc.CreateParagraph().Set(paragraph);

            return doc;
        }

        #endregion Utils

        #region Methods

        public Word ExportFromTemplate<T>(string templateUrl, T wordData) where T : class, new()
        {
            XWPFDocument word = NPOIHelper.GetXWPFDocument(templateUrl);

            ReplacePlaceholders(word, wordData);

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = word.ToBytes()
            };

            return result;
        }

        public Word CreateWord()
        {
            var doc = new XWPFDocument();

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        public Word InsertParagraphs(Word word, List<Paragraph> paragraphs)
        {
            XWPFDocument doc = null;
            using (MemoryStream ms = new MemoryStream(word.WordBytes))
            {
                doc = new XWPFDocument(ms);
            }
            foreach (var p in paragraphs)
            {
                doc = InsertParagraph(doc, p);
            }

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        public Word InsertTables(Word word, List<Table> tables)
        {
            XWPFDocument doc = word.ToNPOI();

            foreach (var t in tables)
            {
                doc = InsertTable(doc, t);
            }

            var result = new Word()
            {
                Type = SolutionEnum.NPOI,
                WordBytes = doc.ToBytes()
            };

            return result;
        }

        public Word CreateFromMasterTable<T>(string templateUrl, IEnumerable<T> datas) where T : class, new()
        {
            var template = NPOIHelper.GetXWPFDocument(templateUrl);

            var result = new Word()
            {
                Type = SolutionEnum.NPOI
            };

            var tables = template.GetTablesEnumerator();

            if (tables == null)
            {
                result.WordBytes = template.ToBytes();
                return result;
            }

            var masterTables = new List<XWPFTable>();

            while (tables.MoveNext())
            {
                masterTables.Add(tables.Current);
            }

            var doc = new XWPFDocument();

            foreach (var data in datas)
            {
                foreach (var masterTable in masterTables)
                {
                    var cloneTable = doc.CreateTable();
                    NPOIHelper.CopyTable(masterTable, cloneTable);
                }

                ReplacePlaceholders(doc, data);
            }

            result.WordBytes = doc.ToBytes();

            return result;
        }

        #endregion Methods

        #region SelfUseMethods

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
        public Word CreateComplexDoc<T>(string templateUrl, T data, int[] tmpTblIndex) where T : class, new()
        {
            XWPFDocument template = NPOIHelper.GetXWPFDocument(templateUrl);

            var result = new Word()
            {
                Type = SolutionEnum.NPOI
            };

            var replaceDict = GeneratePlacehoderDict4ComplexType(data);
            //1:文章段落模板替换
            //2:文章图片替换
            Dictionary<string, IEnumerable<Picture>> pictureReplacements = WordHelper.GetPictureReplacements(data);
            if (replaceDict.TryGetValue(ParagraphKeyName, out List<Dictionary<string, string>> graphReplaces))
            {
                if ((graphReplaces?.Count ?? 0) > 0)
                {
                    var stringReplacements = graphReplaces[0];
                    NPOIHelper.ReplacePlaceholdersInWord(template, stringReplacements, pictureReplacements);
                }
            }

            //3: 模板表格数据替换
            var docTables = template.GetTablesEnumerator();
            var templateTbls = new List<XWPFTable>();
            if (docTables != null && (tmpTblIndex?.Length ?? 0) > 0)
            {
                Array.Sort(tmpTblIndex);
                List<int> sortedTblIndex = tmpTblIndex.ToList();
                int doctblIndex = 0;
                while (docTables.MoveNext())
                {
                    int searchIndex = -1;
                    if (sortedTblIndex.Count > 0) searchIndex = sortedTblIndex.First();
                    if (doctblIndex == searchIndex)
                    {
                        sortedTblIndex.RemoveAt(0);
                        templateTbls.Add(docTables.Current);
                    }
                    doctblIndex++;
                }
            }
            var tableReplaceLists = replaceDict.Where(d => !d.Key.Equals(ParagraphKeyName, StringComparison.CurrentCultureIgnoreCase))
                                    .Select(d => d.Value).ToList();
            if (tableReplaceLists.Count > 0 && templateTbls.Count > 0)
            {
                for (int tIndex = 0; tIndex < tableReplaceLists.Count; tIndex++)
                {//Arrays for table Datas
                 //模板表格与表格占位列表数据位置必须一一对应
                    var tmptbl = templateTbls[tIndex];
                    var cloneRow = tmptbl.GetRow(1); //第一行为表头行，跳过
                    var tmpDoc = new XWPFDocument();
                    var cloneTable = tmpDoc.CreateTable();
                    //在空文档中替换模板表格，最后合并到模板文档中
                    NPOIHelper.CopyTable(tmptbl, cloneTable);

                    var replaceList = tableReplaceLists.ElementAt(tIndex);
                    int maxCol = 0;
                    for (int i = 0; i < replaceList.Count; i++)
                    {
                        var dataRow = replaceList[i];
                        NPOIHelper.ReplacePlaceholdersInWord(tmpDoc, dataRow, null);
                        if ((i + 1) < replaceList.Count)
                        {//最后一行不再添加模板行
                            if (maxCol <= 0)
                                maxCol = cloneTable.Rows.Max(x => x.GetTableCells().Count);
                            NPOIHelper.CopyRowToTable(cloneTable, cloneRow, 1 /*非0即可*/, maxCol);
                        }
                    }
                    //将空文档中替换好数据的表格复制回模板表
                    //NPOIHelper.CopyTable(cloneTable, tmptbl);
                    for (int i = 1; i < cloneTable.Rows.Count; i++)
                    {//首行为表头行，跳过
                        NPOIHelper.CopyRowToTable(tmptbl, cloneTable.Rows[i], 1, maxCol);
                    }
                    tmptbl.RemoveRow(1); //删除模板表格中的占位符行
                }
            }

            result.WordBytes = template.ToBytes();
            return result;
        }

        #endregion SelfUseMethods
    }
}