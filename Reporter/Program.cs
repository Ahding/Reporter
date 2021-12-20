using Reporter.Models;
using Reporter.Tools;
using System.Diagnostics;

const string path = "D:\\Project\\其他";

{
    var parameter = _GetDataParameter(true);

    var exporter = new Report<DataModel>(parameter).GetExportTool();

    var (title, mimetype, extension, dataXlsx) = (parameter.SheetName, exporter.GetMimeType(), exporter.GetExtension(), exporter.Export());

    //Web
    //return System.IO.File(dataXlsx, mimetype, $"{title}.{extension}");

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
/// 取得報表參數_01
/// </summary>
Report<DataModel>.ReportParameterModel _GetDataParameter(bool merge = false)
{
    //取資料
    var source = _GetData();

    #region 欄位合併
    var mergeDataColumnCount = new Dictionary<string, List<NPOIExportTool.MergeDataColumnModel>>();
    var mergeRowCountDic = new Dictionary<string, List<int>>();

    if (merge)
    {
        var sourceGroup = source.GroupBy(x => x.Organization);
        var dataList = new List<DataModel>();

        //欄合併
        var mergeColumnList = new List<NPOIExportTool.MergeDataColumnModel>();
        foreach (var group in sourceGroup)
        {
            var sum = new DataModel
            {
                Organization = group.Key,
                ID = "人數",
                Remark = group.Count().ToString() + "人",
            };
            var result = group.ToList();
            result.Add(sum);
            dataList.AddRange(result);

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
    }
    #endregion

    var title = "簡易測試報表";
    var datas = source.Select(x => (IReport)x).ToList();
    var parameter = new Report<DataModel>.ReportParameterModel
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
        MergeRowCountDic = mergeRowCountDic,
        MergeDataColumnCount = mergeDataColumnCount
    };

    return parameter;
}

/// <summary>
/// 塞資料
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
            Remark = "我的備註很長我的備註很長我的備註很長我的備註很長"
        });
    }

    return datas;
}