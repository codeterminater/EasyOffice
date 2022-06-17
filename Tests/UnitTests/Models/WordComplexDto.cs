using EasyOffice.Models.Word;
using System.Collections.Generic;

namespace UnitTests.Models
{
    public class TableViewModel : BaseTableEle
    {
        /// <summary>
        /// Parent: JszkpdStaticViewModel which is inherit from BaseTableEle
        /// </summary>
        public TableViewModel() : base()
        {
            TotalRoadSum = 0;
            PqiYdlSum = 0;
            PqiLdlSum = 0;
            PqiZdlSum = 0;
            PqiCidlSum = 0;
            PqiChadlSum = 0;
            PQIYLPercentage = 0;
        }

        public double? TotalRoadSum { get; set; }

        /// <summary>
        /// 优等路段里程数
        /// </summary>
        public decimal? PqiYdlSum { get; set; }

        /// <summary>
        /// 良等路段里程数
        /// </summary>
        public decimal? PqiLdlSum { get; set; }

        /// <summary>
        /// 中等路段里程数
        /// </summary>
        public decimal? PqiZdlSum { get; set; }

        /// <summary>
        /// 次等路段里程数
        /// </summary>
        public decimal? PqiCidlSum { get; set; }

        /// <summary>
        /// 差等路段里程数
        /// </summary>
        public decimal? PqiChadlSum { get; set; }

        /// <summary>
        /// PQI优良路率
        /// </summary>
        public double PQIYLPercentage { get; set; }
    }

    public class WordComplexDto
    {
        /// <summary>
        /// 默认数据初始化
        /// </summary>
        public WordComplexDto()
        {
            PQIResult = string.Empty;
            PQITblDatas = new List<TableViewModel>();
        }

        /// <summary>
        /// 检测评定入库里程, 单位公里
        /// </summary>
        public double RealRoadTotalLen { get; set; }

        /// <summary>
        /// 平均PQI
        /// </summary>
        public double PQIAvgPercentage { get; set; }

        /// <summary>
        /// 总体达到PQI: gteq: 大于等于 lt: 小于
        /// </summary>
        public string PQIResult { get; set; }

        //public EasyOffice.Models.Word.Picture PQITablePieImage { get; set; }

        /// <summary>
        /// 一个模板表格对应一个List： PQI评级比例汇总表
        /// </summary>
        public IList<TableViewModel> PQITblDatas { get; set; }
    }
}