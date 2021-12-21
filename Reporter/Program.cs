using System.ComponentModel;
using System.Diagnostics;

const string path = "D:\\Project\\其他";

{
    var parameter = _GetDataParameter();

    var exporter = new Report<DataModel>(parameter).GetExportTool();

    var (title, mimetype, extension, dataXlsx) = (parameter.SheetName, exporter.GetMimeType(), exporter.GetExtension(), exporter.Export());

    //Web
    //return File(dataXlsx, mimetype, $"{title}.{extension}");

    //應用程式
    var fileName = $"{path}\\{title}.{extension}";
    File.WriteAllBytes(fileName, dataXlsx);
    using (var process = new Process())
    {
        process.StartInfo.FileName = fileName;
        process.StartInfo.UseShellExecute = true;
        process.Start();
    }
}

/// <summary>
/// 取得報表參數
/// </summary>
ReportParameterModel _GetDataParameter(bool merge = false)
{
    //取資料
    var source = _GetData();

    #region 欄位合併
    var mergeDataColumnCount = new Dictionary<string, List<NPOIExportTool.MergeDataColumnModel>>();
    var mergeRowCountDic = new Dictionary<string, List<int>>();
    var autoMerge = new ReportParameterModel.AutoMergeModel();

    if (merge)
    {
        var orgGroups = source.GroupBy(x => x.Organization);
        var dataList = new List<DataModel>();

        //插入合計欄位
        var mergeColumnList = new List<NPOIExportTool.MergeDataColumnModel>();
        foreach (var orgGroup in orgGroups)
        {
            var sum = new DataModel
            {
                Organization = orgGroup.Key,
                ID = "人數",
                Remark = orgGroup.Count().ToString() + "人",
            };

            var orgGroupList = orgGroup.ToList();
            orgGroupList.Add(sum);
            dataList.AddRange(orgGroupList);

            //欄合併
            var mergeColumn = new NPOIExportTool.MergeDataColumnModel
            {
                ColumnIdx = 1,
                RowIdx = dataList.Count + 3, //TitleListCount 3 + TableHeader 1 - 以0為起始 1 
                MergeColumnCount = 1,
                MergeCellAlignType = NPOI.SS.UserModel.HorizontalAlignment.Center,
                MergeCellPoints = 16
            };
            mergeColumnList.Add(mergeColumn);
        }

        mergeDataColumnCount.Add(
            nameof(DataModel.ID),
            mergeColumnList
        );

        //列合併
        mergeRowCountDic.Add(
            nameof(DataModel.Organization),
            dataList.GroupBy(x => x.Organization).Select(x => x.Count()).ToList()
        );

        source = dataList;

        autoMerge.MergeRowName = new List<string> { nameof(DataModel.Organization), nameof(DataModel.Remark) };
        autoMerge.MergeColumnName = new List<string> { nameof(DataModel.ID), nameof(DataModel.Remark) };
    }
    #endregion

    var title = "簡易測試報表";
    var datas = source.Select(x => (IReport)x).ToList();
    var parameter = new ReportParameterModel
    {
        Data = datas,
        SheetName = title,
        TitleList = new List<string>
        {
            title,
            "第二行"
        },
        ShowPrintDate = true,
        FooterList = new List<string>
        {
            "這是自訂表尾",
            "第二行"
        },

        //手動
        MergeRowCountDic = mergeRowCountDic,
        MergeDataColumnCount = mergeDataColumnCount,

        //自動
        AutoMerge = autoMerge,
    };

    return parameter;
}

/// <summary>
/// 產生假資料
/// </summary>
List<DataModel> _GetData()
{
    var datas = new List<DataModel>();

    for (int i = 1; i <= 100; i++)
    {
        datas.Add(new DataModel
        {
            ID = $"{i}",
            Organization = $"{(i - 1) / 10 + 1}組",
            Name = $"第{i}位",
            Remark = "我的備註很長我的備註很長我的備註很長我的備註很長",
            //Remark2 = "備註2"
        });
    }

    return datas;
}

/// <summary>
/// 匯出模型範例
/// </summary>
public class DataModel : IReport
{
    /// <summary>單位</summary>
    [ColumnWidth(ColumnWidth = 15)]
    [Description("單位")]
    public string Organization { get; set; }

    /// <summary>ID</summary>
    [ColumnWidth(ColumnWidth = 15)]
    [Description("ID")]
    public string ID { get; set; }


    /// <summary>姓名</summary>
    [ColumnWidth(ColumnWidth = 15)]
    [Description("姓名")]
    public string Name { get; set; }

    /// <summary>備註</summary>
    [ColumnWidth(ColumnWidth = 80)]
    [Description("備註")]
    public string Remark { get; set; }
    /// <summary>備註2</summary>
    //[ColumnWidth(ColumnWidth = 100)]
    //[Description("備註2")]
    //public string Remark2 { get; set; }
}