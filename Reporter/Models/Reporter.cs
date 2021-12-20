using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;
using Reporter.Tools;
using static Reporter.Tools.NPOIExportTool;
using Reporter.Attributes;

namespace Reporter.Models
{

    /// <summary>
    /// 通用匯出報表產生器
    /// </summary>
    public class Report<T> : ReportSheetCreater where T : new()
    {
        private List<string> _TitleList { get; set; }
        private List<string> _FooterList { get; set; }
        private Dictionary<string, List<int>> _MergeRowCountDic { get; set; }
        private Dictionary<string, List<MergeDataColumnModel>> _MergeDataColumnCount { get; set; }

        private bool _ShowEmptyIfZero { get; set; }
        private bool _ShowPrintDate { get; set; }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="parameter"></param>
        public Report(ReportParameterModel parameter) : base(parameter.Exporter ?? new NPOIExportTool(), parameter.SheetName, parameter.Data)
        {
            _TitleList = parameter.TitleList;
            _FooterList = parameter.FooterList;
            _ShowEmptyIfZero = parameter.ShowEmptyIfZero;
            _MergeRowCountDic = parameter.MergeRowCountDic ?? new Dictionary<string, List<int>>();
            _MergeDataColumnCount = parameter.MergeDataColumnCount ?? new Dictionary<string, List<MergeDataColumnModel>>();
            _ShowPrintDate = parameter.ShowPrintDate;

            BindAllData();
        }

        /// <summary>
        /// 取得報表匯出工具
        /// </summary>
        /// <returns></returns>
        public NPOIExportTool GetExportTool()
        {
            return this.Exporter;
        }

        /// <summary>
        /// 設定表身資料
        /// </summary>
        /// <returns></returns>
        protected override List<NPOIExportTool.BindModel> SetupTableBinding()
        {

            var pointNumbers = new List<string> {
                typeof(decimal).ToString(),
                typeof(float).ToString(),
                typeof(double).ToString()
            };
            var intNumber = new List<string> {
                typeof(int).ToString(),
            };
            var properties = (new T()).GetType().GetProperties();
            var dataItem = properties.Select(x => new {
                x.Name,
                Attrs = x.GetCustomAttributes(typeof(DescriptionAttribute), true),
                ColumnAttrs = x.GetCustomAttributes(typeof(ColumnWidthAttribute), true),
                x.PropertyType,
            }).Select(x => new _CustomizedColumnModel
            {
                ColumnAttributeName = x.Name,
                ColumnName = x.Attrs != null && x.Attrs.Length > 0 ? ((DescriptionAttribute)x.Attrs[0]).Description : x.Name,
                ColumnWidth = x.ColumnAttrs != null && x.ColumnAttrs.Length > 0 ? ((ColumnWidthAttribute)x.ColumnAttrs[0]).ColumnWidth : (int?)null,
                IsInteger = x.PropertyType != null ? intNumber.Contains(x.PropertyType.ToString()) : false,
                IsPointNumber = x.PropertyType != null ? pointNumbers.Contains(x.PropertyType.ToString()) : false,
            }).ToList();

            var defaultColumnWidth = 30;
            var bindList = new List<NPOIExportTool.BindModel>();

            foreach (var item in dataItem)
            {
                var bind = new NPOIExportTool.BindModel
                {
                    HeadName = item.ColumnName,
                    DataName = new List<string> {
                           item.ColumnAttributeName,
                        },
                    ColumnWidth = item.ColumnWidth ?? defaultColumnWidth,
                    ShowEmptyIfZero = _ShowEmptyIfZero,

                    MergeDataRowCount = _MergeRowCountDic.ContainsKey(item.ColumnAttributeName) ?
                    _MergeRowCountDic[item.ColumnAttributeName] ?? new List<int>() :
                    new List<int>(),

                    MergeDataColumnCount = _MergeDataColumnCount.ContainsKey(item.ColumnAttributeName) ?
                    _MergeDataColumnCount[item.ColumnAttributeName] ?? new List<MergeDataColumnModel>() :
                    new List<MergeDataColumnModel>(),
                };

                if (item.IsInteger || item.IsPointNumber)
                    bind.ApplyCellAlignmentForColumn = HorizontalAlignment.Right;
                if (item.IsPointNumber)
                    bind.DataWithPoints = 2;

                bindList.Add(bind);
            }

            return bindList;
        }
        /// <summary>
        /// 設定表頭
        /// </summary>
        /// <returns></returns>
        protected override List<NPOIExportTool.TitleModel> SetupTitle()
        {

            var titleList = new List<NPOIExportTool.TitleModel>();

            _TitleList = _TitleList ?? new List<string>();
            foreach (var title in _TitleList)
            {
                titleList.Add(new NPOIExportTool.TitleModel
                {
                    Content = new List<NPOIExportTool.TitleContentModel> {
                    new NPOIExportTool.TitleContentModel{
                            Text = title,
                            MergeColumnCount = TotalColumnCount - 1,
                        },
                    },
                    RowHeightInPoint = 27
                });
            }

            if (_ShowPrintDate)
            {
                var printerCellStyle = Exporter.CloneInitTableCellStyleForString();
                printerCellStyle.Alignment = HorizontalAlignment.Right;
                titleList.Add(new NPOIExportTool.TitleModel
                {

                    Content = new List<NPOIExportTool.TitleContentModel> {
                        new NPOIExportTool.TitleContentModel{
                            Text = string.Format("列印日期：{0}年{1:MM月dd日}", DateTime.Now.Year - 1911, DateTime.Now),
                            MergeColumnCount = TotalColumnCount-1,
                            MergeRowCount = 0,
                            ApplyCellStyle = printerCellStyle,
                        },
                    },
                    RowHeightInPoint = 21
                });
            }

            return titleList;
        }

        /// <summary>
        /// 設定表尾
        /// </summary>
        /// <returns></returns>
        protected override List<NPOIExportTool.TitleModel> SetupFooter()
        {
            
            var footerList = new List<NPOIExportTool.TitleModel>();

            _FooterList = _FooterList ?? new List<string>();
            foreach (var footer in _FooterList)
            {
                footerList.Add(new NPOIExportTool.TitleModel
                {
                    Content = new List<NPOIExportTool.TitleContentModel> {
                    new NPOIExportTool.TitleContentModel{
                            Text = footer,
                            MergeColumnCount = TotalColumnCount - 1,
                        },
                    }
                });
            }

            return footerList;
        }


        private class _CustomizedColumnModel
        {
            public int? ColumnWidth { get; set; }
            public string ColumnAttributeName { get; set; }
            public string ColumnName { get; set; }
            public bool IsInteger { get; set; }
            public bool IsPointNumber { get; set; }
            public List<int> MergeRowCount { get; set; }
        }

        /// <summary>
        /// 匯出參數模型
        /// </summary>
        public class ReportParameterModel
        {
            /// <summary>匯出服務</summary>
            public NPOIExportTool Exporter { get; set; }
            /// <summary>分頁名稱</summary>
            public string SheetName { get; set; }
            /// <summary>匯出資料</summary>
            public List<IReport> Data { get; set; }
            /// <summary>標題清單</summary>
            public List<string> TitleList { get; set; }
            /// <summary>顯示列印日期</summary>
            public bool ShowPrintDate { get; set; }
            /// <summary>表尾清單</summary>
            public List<string> FooterList { get; set; }
            public bool ShowEmptyIfZero { get; set; }
            /// <summary>合併列</summary>
            public Dictionary<string, List<int>> MergeRowCountDic { get; set; }
            /// <summary>合併欄</summary>
            public Dictionary<string, List<MergeDataColumnModel>> MergeDataColumnCount { get; set; }

        }
    }
}