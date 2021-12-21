namespace Reporter.Tools
{
    #region 工具
    /// <summary>
    /// NPOI 匯出工具
    /// </summary>
    public class NPOIExportTool
    {
        /// <summary>
        /// 是否要啟用自適應寬度
        /// </summary>
        public bool IsAutoColumn = true;

        /// <summary>
        /// 預設的欄位寬度
        /// </summary>
        public int DefaultColumnSize = 20;

        #region 基本顏色設定
        private XSSFColor _blackColor = new XSSFColor(new byte[] { 0, 0, 0 });
        private XSSFColor _gray = new XSSFColor(new byte[] { 222, 222, 222 });
        private XSSFColor _veryLightPinkColor = new XSSFColor(new byte[] { 255, 232, 232 });
        private XSSFColor _white = new XSSFColor(new byte[] { 255, 255, 255 });
        #endregion

        /// <summary> 預設字體大小 </summary>
        private Dictionary<TableLayerEnum, short> _DefaultFontSize { get; set; }
        /// <summary> 活頁簿 </summary>
        protected IWorkbook _Workbook = new XSSFWorkbook();
        /// <summary> 預設標題列儲存格格式 </summary>
        private XSSFCellStyle _TitleCellStyle { get; set; }
        /// <summary> 預設表格表頭列儲存格格式 </summary>
        private XSSFCellStyle _HeadCellStyle { get; set; }
        /// <summary> 預設表格資料列儲存格格式 </summary>
        private Dictionary<int, XSSFCellStyle> _DataCellStyle { get; set; }
        /// <summary> 預設表尾列儲存格格式 </summary>
        private XSSFCellStyle _FooterCellStyle { get; set; }
        /// <summary>分頁清單</summary>
        protected List<SheetInfoModel<IReport>> _SheetList { get; set; }

        #region 初始化
        /// <summary>
        /// 建構子
        /// </summary>
        public NPOIExportTool()
        {
            Reset();

        }

        /// <summary>
        /// 重新設定匯出工具
        /// </summary>
        public void Reset()
        {
            this._Workbook = new XSSFWorkbook();
            this._SheetList = new List<SheetInfoModel<IReport>>();
            // 初始化 預設資料格格式
            _InitDefaultFont(20);
            _InitTitleCellStyle();
            _InitHeadCellStyle();
            _InitDateCellStyle();
            _InitFooterCellStyle();
        }
        /// <summary>
        /// 加入分頁
        /// </summary>
        /// <param name="sheet"></param>
        public void AddSheet(SheetInfoModel<IReport> sheet)
        {
            this._SheetList.Add(sheet);
        }
        /// <summary>
        /// 加入分頁清單
        /// </summary>
        /// <param name="sheetList"></param>
        public void AddSheetList(List<SheetInfoModel<IReport>> sheetList)
        {
            this._SheetList.AddRange(sheetList);
        }
        /// <summary>初始化標題列 預設資料格格式</summary>
        private void _InitTitleCellStyle()
        {
            var workbook = this._Workbook;
            var titleStyle = this._TitleCellStyle;
            // 字型
            XSSFFont titleFont = (XSSFFont)workbook.CreateFont();
            titleFont.FontName = "標楷體";
            titleFont.FontHeightInPoints = this._DefaultFontSize[TableLayerEnum.Title];
            titleFont.IsBold = true;
            this._TitleCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            this._TitleCellStyle.SetFont(titleFont);
            this._TitleCellStyle.SetBorderColor(BorderSide.TOP, _blackColor);
            this._TitleCellStyle.SetBorderColor(BorderSide.RIGHT, _blackColor);
            this._TitleCellStyle.SetBorderColor(BorderSide.BOTTOM, _blackColor);
            this._TitleCellStyle.SetBorderColor(BorderSide.LEFT, _blackColor);
            this._TitleCellStyle.Alignment = HorizontalAlignment.Center;
            this._TitleCellStyle.VerticalAlignment = VerticalAlignment.Center;
            this._TitleCellStyle.BorderTop = BorderStyle.Thin;
            this._TitleCellStyle.BorderRight = BorderStyle.Thin;
            this._TitleCellStyle.BorderBottom = BorderStyle.Thin;
            this._TitleCellStyle.BorderLeft = BorderStyle.Thin;
            this._TitleCellStyle.WrapText = true;
        }
        /// <summary>初始化表頭列 預設資料格格式</summary>
        private void _InitHeadCellStyle()
        {
            var workbook = this._Workbook;
            var titleStyle = this._TitleCellStyle;


            // 字型
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.FontName = "標楷體";
            font.FontHeightInPoints = this._DefaultFontSize[TableLayerEnum.Head];
            font.IsBold = true;

            this._HeadCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            this._HeadCellStyle.CloneStyleFrom(titleStyle);
            this._HeadCellStyle.SetFont(font);
            this._HeadCellStyle.Alignment = HorizontalAlignment.Center;
            this._HeadCellStyle.SetFillForegroundColor(new XSSFColor(new byte[] { 237, 245, 217 }));
            this._HeadCellStyle.FillPattern = FillPattern.SolidForeground;
        }
        /// <summary>初始化資料列 預設資料格格式</summary>
        private void _InitDateCellStyle()
        {
            var workbook = this._Workbook;
            var titleStyle = this._TitleCellStyle;
            this._DataCellStyle = new Dictionary<int, XSSFCellStyle>();

            // 字型
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.FontName = "標楷體";
            font.FontHeightInPoints = this._DefaultFontSize[TableLayerEnum.Data];
            font.IsBold = false;
            // 只有整數的數字欄位格式欄位
            var numberIntCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            numberIntCellStyle.CloneStyleFrom(titleStyle);
            numberIntCellStyle.SetFont(font);
            numberIntCellStyle.SetFillForegroundColor(_white);
            numberIntCellStyle.FillPattern = FillPattern.SolidForeground;
            numberIntCellStyle.Alignment = HorizontalAlignment.Center;
            numberIntCellStyle.SetDataFormat(HSSFDataFormat.GetBuiltinFormat("#,##0"));
            this._DataCellStyle.Add((int)DataCellDataTypeEnum.NumberInt, numberIntCellStyle);
            // 有小數的數字欄位格式欄位
            var numberWithPointCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            numberWithPointCellStyle.CloneStyleFrom(titleStyle);
            numberWithPointCellStyle.SetFont(font);
            numberWithPointCellStyle.SetFillForegroundColor(_white);
            numberWithPointCellStyle.FillPattern = FillPattern.SolidForeground;
            numberWithPointCellStyle.Alignment = HorizontalAlignment.Center;
            numberWithPointCellStyle.SetDataFormat(HSSFDataFormat.GetBuiltinFormat("#,##0.00"));
            this._DataCellStyle.Add((int)DataCellDataTypeEnum.NumberWithPoint, numberWithPointCellStyle);
            // 字串的欄位格式欄位
            var stringCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            stringCellStyle.CloneStyleFrom(titleStyle);
            stringCellStyle.SetFont(font);
            stringCellStyle.SetFillForegroundColor(_white);
            stringCellStyle.FillPattern = FillPattern.SolidForeground;
            stringCellStyle.Alignment = HorizontalAlignment.Center;
            this._DataCellStyle.Add((int)DataCellDataTypeEnum.String, stringCellStyle);

        }
        /// <summary>初始化表尾列 預設資料格格式</summary>
        private void _InitFooterCellStyle()
        {
            var workbook = this._Workbook;


            // 字型
            XSSFFont font = (XSSFFont)workbook.CreateFont();
            font.FontName = "標楷體";
            font.FontHeightInPoints = this._DefaultFontSize[TableLayerEnum.Footer];
            font.IsBold = false;


            this._FooterCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            this._FooterCellStyle.Alignment = HorizontalAlignment.Left;
            this._FooterCellStyle.VerticalAlignment = VerticalAlignment.Center;
            this._FooterCellStyle.BorderTop = BorderStyle.None;
            this._FooterCellStyle.BorderRight = BorderStyle.None;
            this._FooterCellStyle.BorderBottom = BorderStyle.None;
            this._FooterCellStyle.BorderLeft = BorderStyle.None;
            this._FooterCellStyle.SetFont(font);
            this._FooterCellStyle.Alignment = HorizontalAlignment.Left;

        }
        /// <summary>初始化表尾列 預設字型</summary>
        private void _InitDefaultFont(int maxFontSize)
        {
            this._DefaultFontSize = new Dictionary<TableLayerEnum, short> {
                    { TableLayerEnum.Title, (short)maxFontSize},
                    { TableLayerEnum.Head, (short)(maxFontSize - 4)},
                    { TableLayerEnum.Data, (short)(maxFontSize - 4)},
                    { TableLayerEnum.Footer, (short)(maxFontSize - 4)},
                };

        }

        #endregion

        #region 執行匯出

        /// <summary>
        /// 執行匯出
        /// </summary>
        /// <returns>byte array</returns>
        public virtual byte[] Export()
        {
            var today = DateTime.Now;
            // 產生 xlsx 檔案
            IWorkbook workbook = this._Workbook;

            foreach (var sheetInfo in this._SheetList)
            {
                // 產生試算表
                XSSFSheet sheet = (XSSFSheet)workbook.CreateSheet(sheetInfo.Name);
                sheet.PrintSetup.ValidSettings = true;
                sheet.PrintSetup.Landscape = sheetInfo.PrintLandScape;
                if (sheetInfo.PrintPageSize > 0)
                {
                    sheet.PrintSetup.PaperSize = sheetInfo.PrintPageSize;
                }
                sheet.FitToPage = true;
                sheet.PrintSetup.FitHeight = 0;
                sheet.PrintSetup.FitWidth = 1;

                var dataRowIndex = 0;

                _FullFillTitle(sheet, sheetInfo.GetTitleList(), false, ref dataRowIndex);
                _FullFillTable(sheet, sheetInfo.GetDataBindingList(), sheetInfo.GetDataList(), ref dataRowIndex);

                dataRowIndex++;
                _FullFillFooter(sheet, sheetInfo.GetDataFooterList(), ref dataRowIndex);
            }


            // 計算所有公式欄位的值
            XSSFFormulaEvaluator.EvaluateAllFormulaCells(workbook);

            // 輸出 xlsx 檔案
            var memoryStream = new MemoryStream();
            workbook.Write(memoryStream);


            var byteResult = memoryStream.ToArray();
            if (_CusFontDic != null)
                _CusFontDic.Clear();
            return byteResult;

        }
        /// <summary>
        /// 填寫標題資訊
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataList"></param>
        /// <param name="isFooter"></param>
        /// <param name="dataRowIndex"></param>
        protected void _FullFillTitle(XSSFSheet sheet, List<TitleModel> dataList, bool isFooter, ref int dataRowIndex)
        {
            var defaultCellStyle = isFooter ? this._FooterCellStyle : this._TitleCellStyle;
            var titleRows = new Dictionary<int, XSSFRow>();
            var currentRow = dataRowIndex;
            foreach (var title in dataList)
            {
                var rowHeightInPoint = title.RowHeightInPoint;
                var nowRow = currentRow;
                var nowCol = 0;
                foreach (var titleContent in title.Content)
                {
                    int startRow = nowRow, endRow = nowRow;
                    int startColumn = nowCol, endColumn = nowCol;

                    if (titleContent.MergeColumnCount.HasValue || titleContent.MergeRowCount.HasValue)
                    {
                        if (titleContent.MergeColumnCount.HasValue)
                        {
                            startColumn = nowCol;
                            endColumn = nowCol + Math.Abs(titleContent.MergeColumnCount.Value);
                        }
                        if (titleContent.MergeRowCount.HasValue)
                        {
                            startRow = nowRow;
                            endRow = nowRow + Math.Abs(titleContent.MergeRowCount.Value);
                        }

                        if (titleContent.MergeColumnCount.Value > 0 || (titleContent.MergeRowCount.HasValue && titleContent.MergeRowCount.Value > 0))
                        {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startColumn, endColumn));
                        }

                    }

                    var initialRow = startRow;
                    var initialColumn = startColumn;
                    var applyCellStyle = titleContent.ApplyCellStyle != null ? titleContent.ApplyCellStyle : defaultCellStyle;
                    while (startRow <= endRow)
                    {
                        if (!titleRows.ContainsKey(startRow))
                        {
                            var thisRow = (XSSFRow)sheet.CreateRow(startRow);
                            if (rowHeightInPoint != -1) { thisRow.HeightInPoints = rowHeightInPoint; }
                            titleRows.Add(startRow, thisRow);
                        }
                        var dataRow = titleRows[startRow];

                        while (startColumn <= endColumn)
                        {
                            _CreateCell(dataRow, startColumn++, CellType.String, applyCellStyle).SetCellValue("");

                        }
                        startRow++;
                        startColumn = initialColumn;
                    }
                    _CreateCell(titleRows[initialRow], initialColumn, CellType.String, applyCellStyle).SetCellValue(titleContent.Text);
                    nowCol = endColumn + 1;

                }
                currentRow++;
                dataRowIndex++;
            }


        }


        /// <summary>
        /// 填寫資料表資訊
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataBinding"></param>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex"></param>
        protected void _FullFillTable(XSSFSheet sheet, List<BindModel> dataBinding, List<IReport> dataList, ref int dataRowIndex)
        {


            var maxRowCount = dataBinding.Max(x => x.MaxRowCount);

            #region 資料表標題
            var headRowDic = new Dictionary<int, XSSFRow>();
            _CreateHeadRow(sheet, maxRowCount, dataBinding, headRowDic, dataRowIndex);
            dataRowIndex += maxRowCount;
            #endregion

            #region 資料表內容
            var dataCount = dataList.Count();
            if (dataCount > 0)
            {
                var dataNameList = new List<DataNameInfoModel>();
                _BuildDataNameOfDataBinding(dataBinding, dataNameList);
                var dataNameMergeInfo = new Dictionary<string, DataNameMergeInfoModel>();
                int dataNameCounter = 0;
                foreach (var dataNameInfo in dataNameList)
                {
                    var dataName = string.Format("{0}___{1}", dataNameCounter, dataNameInfo.DataName);
                    dataNameMergeInfo.Add(dataName, new DataNameMergeInfoModel
                    {
                        State = false,
                        Counter = 0,
                        Pointer = 0
                    });
                    dataNameInfo.DataName = dataName;
                    dataNameCounter++;
                }

                foreach (var item in dataList)
                {
                    XSSFRow dataRow = (XSSFRow)sheet.CreateRow(dataRowIndex);
                    var data = item.GetType().GetProperties().ToDictionary(
                        x => x.Name,
                        x => new {
                            value = x.GetValue(item, null),
                            datatype = x.PropertyType
                        }
                    );
                    int col = 0;
                    int mergeCol = 0;
                    object firstMergeCellVal = null;
                    HorizontalAlignment firstMergeCellAlignType = 0;
                    short firstMergeCellPoints = 0;
                    foreach (var dataNameInfo in dataNameList)
                    {
                        var printData = true;
                        var dataName = dataNameInfo.DataName;
                        var pureDataName = (dataName.Split(new[] { "___" }, StringSplitOptions.RemoveEmptyEntries))[1];
                        var mergeInfo = dataNameMergeInfo[dataName];
                        if (mergeInfo.Pointer < dataNameInfo.NeedMergeRowCount.Count())
                        {
                            if (!mergeInfo.State && mergeInfo.Counter < dataNameInfo.NeedMergeRowCount[mergeInfo.Pointer])
                            {
                                mergeInfo.State = true;
                                dataNameInfo.NeedMergeRowCount[mergeInfo.Pointer] = dataNameInfo.NeedMergeRowCount[mergeInfo.Pointer] - 1;
                                mergeInfo.Counter = dataNameInfo.NeedMergeRowCount[mergeInfo.Pointer];
                                if (mergeInfo.Counter > 0)
                                {
                                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(dataRowIndex, dataRowIndex + mergeInfo.Counter, col, col));
                                }
                                // do Merge
                            }
                            else if (mergeInfo.Counter > 0)
                            {
                                printData = false;
                            }

                            if (mergeInfo.Counter <= 0)
                            {
                                mergeInfo.Pointer++;
                                mergeInfo.State = false;

                            }
                            mergeInfo.Counter--;
                        }

                        int rowIdx = dataRowIndex;
                        if (dataNameInfo.NeedMergeColumnCount.Count > 0 && dataNameInfo.NeedMergeColumnCount.Any(x => x.RowIdx == rowIdx && x.ColumnIdx == col))
                        {
                            var mergeColumnInfo = dataNameInfo.NeedMergeColumnCount.FirstOrDefault(x => x.RowIdx == rowIdx && x.ColumnIdx == col);
                            int needMergeColumnCount = mergeColumnInfo.MergeColumnCount;
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(dataRowIndex, dataRowIndex, col, col + needMergeColumnCount));
                            mergeCol = needMergeColumnCount + 1;
                            firstMergeCellVal = data[pureDataName].value;
                            firstMergeCellAlignType = mergeColumnInfo.MergeCellAlignType;
                            firstMergeCellPoints = mergeColumnInfo.MergeCellPoints;
                        }
                        if (mergeCol > 0)
                        {
                            var _mergeCellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.String];
                            if (firstMergeCellAlignType > 0)
                                _mergeCellStyle.Alignment = firstMergeCellAlignType;
                            if (firstMergeCellPoints > 0)
                            {
                                var _firstFont = CreateNewFont();
                                _firstFont.FontHeightInPoints = firstMergeCellPoints;
                                _mergeCellStyle.SetFont(_firstFont);
                            }

                            _CreateCell(dataRow, col++, CellType.String, _mergeCellStyle).SetCellValue((string)firstMergeCellVal);
                            mergeCol--;
                        }
                        else
                        {
                            if (printData)
                            {
                                if (dataNameInfo.ApplyCellStyleForColumn != null)
                                {
                                    var cellStyle = dataNameInfo.ApplyCellStyleForColumn;
                                    _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                    _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue((string)data[pureDataName].value);
                                }
                                else if (data[pureDataName].datatype == typeof(string))
                                {
                                    var value = (string)data[pureDataName].value;
                                    bool tryParseWithNum = double.TryParse(value, out double valueWithNum);
                                    if (tryParseWithNum && dataNameInfo.IsStringWithNumber)
                                    {
                                        if (dataNameInfo.ShowEmptyIfZero && valueWithNum == 0)
                                        {
                                            var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.String];
                                            _SetDataCellOtherParameters(cellStyle, dataNameInfo, true);
                                            _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(dataNameInfo.EmptyString);
                                        }
                                        else if (valueWithNum == (int)valueWithNum && ((dataNameInfo.DataWithPoints ?? 0) == 0))
                                        {
                                            var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                            _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                            _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(valueWithNum);
                                        }
                                        else
                                        {
                                            var points = dataNameInfo.DataWithPoints.HasValue ? dataNameInfo.DataWithPoints.Value : 2;
                                            valueWithNum = Math.Round(valueWithNum, points, MidpointRounding.AwayFromZero);
                                            bool IsForInt = ((points == 0) && (valueWithNum % 1 == 0));
                                            var cellStyle = IsForInt ? this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt] : CloneInitTableCellStyleForNumberPoint();
                                            _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                            _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(valueWithNum);
                                        }
                                    }
                                    else
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.String];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo, true);
                                        _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(value);
                                    }
                                }
                                else if (data[pureDataName].datatype == typeof(double))
                                {
                                    var value = (double)data[pureDataName].value;
                                    if (dataNameInfo.ShowEmptyIfZero && value == 0)
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(dataNameInfo.EmptyString);
                                    }
                                    else if (value == (int)value && ((dataNameInfo.DataWithPoints ?? 0) == 0))
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }
                                    else
                                    {
                                        var points = dataNameInfo.DataWithPoints.HasValue ? dataNameInfo.DataWithPoints.Value : 2;
                                        value = Math.Round(value, points, MidpointRounding.AwayFromZero);
                                        bool IsForInt = ((points == 0) && (value % 1 == 0));
                                        var cellStyle = IsForInt ? this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt] : CloneInitTableCellStyleForNumberPoint();
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }
                                }
                                else if (data[pureDataName].datatype == typeof(decimal))
                                {
                                    var value = Convert.ToDouble((decimal)data[pureDataName].value);
                                    if (dataNameInfo.ShowEmptyIfZero && value == 0)
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(dataNameInfo.EmptyString);
                                    }
                                    else if (value == (int)value && ((dataNameInfo.DataWithPoints ?? 0) == 0))
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }
                                    else
                                    {
                                        var points = dataNameInfo.DataWithPoints.HasValue ? dataNameInfo.DataWithPoints.Value : 2;
                                        value = Math.Round(value, points, MidpointRounding.AwayFromZero);
                                        bool IsForInt = ((points == 0) && (value % 1 == 0));
                                        var cellStyle = IsForInt ? this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt] : CloneInitTableCellStyleForNumberPoint();
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }

                                }
                                else if (data[pureDataName].datatype == typeof(float))
                                {
                                    var value = (float)data[pureDataName].value;
                                    if (dataNameInfo.ShowEmptyIfZero && value == 0)
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(dataNameInfo.EmptyString);
                                    }
                                    else if (value == (int)value && ((dataNameInfo.DataWithPoints ?? 0) == 0))
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }
                                    else
                                    {
                                        var points = dataNameInfo.DataWithPoints.HasValue ? dataNameInfo.DataWithPoints.Value : 2;
                                        var doubleValue = Math.Round(value, points, MidpointRounding.AwayFromZero);
                                        bool IsForInt = ((points == 0) && (doubleValue % 1 == 0));
                                        var cellStyle = IsForInt ? this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt] : CloneInitTableCellStyleForNumberPoint();
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(doubleValue);
                                    }
                                }
                                else if (data[pureDataName].datatype == typeof(int))
                                {
                                    var value = (int)data[pureDataName].value;
                                    if (dataNameInfo.ShowEmptyIfZero && value == 0)
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.String, cellStyle).SetCellValue(dataNameInfo.EmptyString);
                                    }
                                    else
                                    {
                                        var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt];
                                        _SetDataCellOtherParameters(cellStyle, dataNameInfo);
                                        _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue(value);
                                    }
                                }
                            }
                            else
                            {
                                var cellStyle = this._DataCellStyle[(int)DataCellDataTypeEnum.String];
                                if (dataNameInfo.ApplyCellAlignmentForColumn > 0)
                                {
                                    cellStyle.Alignment = dataNameInfo.ApplyCellAlignmentForColumn;
                                }
                                _CreateCell(dataRow, col++, CellType.Numeric, cellStyle).SetCellValue("");
                            }
                        }

                    }


                    ++dataRowIndex;
                }

            }
            #endregion

        }

        /// <summary>
        /// 客制化字形字典集
        /// </summary>
        protected Dictionary<int, XSSFFont> _CusFontDic { get; set; }
        /// <summary>
        /// 設定資料格式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="dataNameInfo"></param>
        /// <param name="isStringExceptAlignment"></param>
        private void _SetDataCellOtherParameters(XSSFCellStyle cellStyle, DataNameInfoModel dataNameInfo, bool isStringExceptAlignment = false)
        {
            //對齊方式
            if (dataNameInfo.ApplyCellAlignmentForColumn > 0 && !isStringExceptAlignment)
            {
                cellStyle.Alignment = dataNameInfo.ApplyCellAlignmentForColumn;
            }
            //字體大小
            if (dataNameInfo.DataFontPoints > 0)
            {
                if (_CusFontDic == null)
                    _CusFontDic = new Dictionary<int, XSSFFont>();

                if (!_CusFontDic.ContainsKey(dataNameInfo.DataFontPoints))
                {
                    var cellfont = CreateNewFont();
                    cellfont.FontHeightInPoints = dataNameInfo.DataFontPoints;
                    _CusFontDic.Add(dataNameInfo.DataFontPoints, cellfont);
                }

                cellStyle.SetFont(_CusFontDic[dataNameInfo.DataFontPoints]);
            }
        }
        /// <summary>
        /// 填寫表頭資訊
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="maxRowCount"></param>
        /// <param name="dataBinding"></param>
        /// <param name="rowDic"></param>
        /// <param name="dataRowIndex"></param>
        private void _CreateHeadRow(XSSFSheet sheet, int maxRowCount, List<BindModel> dataBinding, Dictionary<int, XSSFRow> rowDic, int dataRowIndex)
        {
            if (!rowDic.ContainsKey(dataRowIndex))
            {
                rowDic.Add(dataRowIndex, (XSSFRow)sheet.CreateRow(dataRowIndex));
            }
            var dataHeadRow = rowDic[dataRowIndex];
            int nowColumnIndex = 0, colHead = 0;
            bool hasNextRow = false;
            var nextRowDataBinding = new List<BindModel>();
            foreach (var item in dataBinding)
            {

                var dataColumnCount = 0; var thisHasNextRow = false;
                colHead = nowColumnIndex;
                if (item.SubDataBinding.Any())
                {
                    hasNextRow = true;
                    thisHasNextRow = true;
                    nextRowDataBinding.AddRange(item.SubDataBinding);
                }
                else
                {
                    nextRowDataBinding.Add(new BindModel
                    {
                        RowNo = -1,
                        ColumsCount = item.ColumsCount,
                        ColumnWidth = item.ColumnWidth,
                    });

                }

                dataColumnCount = item.ColumsCount;

                if (item.RowNo > -1)
                {
                    var mergeRowCount = thisHasNextRow ? 0 : ((maxRowCount - 1) >= 0 ? (maxRowCount - 1) : 0);
                    if (!thisHasNextRow || (dataColumnCount > 1 && thisHasNextRow))
                    {
                        var startRow = dataRowIndex;
                        var endRow = dataRowIndex + mergeRowCount;
                        var startColumn = nowColumnIndex;
                        var endColumn = nowColumnIndex + dataColumnCount - 1;
                        if (endRow > startRow || endColumn > startColumn)
                        {
                            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRow, endRow, startColumn, endColumn));
                        }

                    }

                }
                var isFirst = true;
                for (var i = 0; i < dataColumnCount; i++)
                {
                    var applyCellStyle = item.ApplyCellStyle == null ? CloneInitHeadCellStyle() : item.ApplyCellStyle;
                    var column = DefaultColumnSize;
                    if (item.ColumnWidth.HasValue)
                    {
                        column = item.ColumnWidth.Value;
                    }
                    sheet.SetColumnWidth(colHead, column * 256);
                    if (item.HeadFontPoints > 0)
                    {
                        var headfont = applyCellStyle.GetFont();
                        headfont.FontHeightInPoints = item.HeadFontPoints;
                        applyCellStyle.SetFont(headfont);
                    }

                    if (isFirst && item.RowNo > -1)
                    {
                        _CreateCell(dataHeadRow, colHead++, CellType.String, applyCellStyle).SetCellValue(item.HeadName);
                        isFirst = false;
                    }
                    else
                    {
                        _CreateCell(dataHeadRow, colHead++, CellType.String, applyCellStyle).SetCellValue("");
                    }

                }

                nowColumnIndex += dataColumnCount;


            }
            if (hasNextRow)
            {
                _CreateHeadRow(sheet, maxRowCount - 1, nextRowDataBinding, rowDic, dataRowIndex + 1);
            }
        }

        /// <summary>
        /// 建立表身的填寫資訊
        /// </summary>
        /// <param name="dataBinding">輸入：表頭：表身關聯表</param>
        /// <param name="dataNameList">輸出：List(欲填寫於表身的的屬性名稱)</param>
        private void _BuildDataNameOfDataBinding(List<BindModel> dataBinding, List<DataNameInfoModel> dataNameList)
        {
            foreach (var item in dataBinding)
            {
                if (item.SubDataBinding.Any())
                {
                    _BuildDataNameOfDataBinding(item.SubDataBinding, dataNameList);
                }
                else
                {
                    foreach (var dataName in item.DataName)
                    {
                        dataNameList.Add(new DataNameInfoModel
                        {
                            DataName = dataName,
                            NeedMergeRowCount = item.MergeDataRowCount,
                            ShowEmptyIfZero = item.ShowEmptyIfZero,
                            EmptyString = item.EmptyString,
                            ApplyCellStyleForColumn = item.ApplyCellStyleForColumn,
                            NeedMergeColumnCount = item.MergeDataColumnCount,
                            ApplyCellAlignmentForColumn = item.ApplyCellAlignmentForColumn,
                            DataFontPoints = item.DataFontPoints,
                            DataWithPoints = item.DataWithPoints,
                            IsStringWithNumber = item.IsStringWithNumber
                        });
                    }

                }
            }
        }


        /// <summary>
        /// 填寫表尾資訊
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="dataList"></param>
        /// <param name="dataRowIndex"></param>

        protected void _FullFillFooter(XSSFSheet sheet, List<TitleModel> dataList, ref int dataRowIndex)
        {

            _FullFillTitle(sheet, dataList, true, ref dataRowIndex);
        }

        /// <summary>
        /// 複製一筆預設的 (標題)儲存格格式
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitCellStyle()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._TitleCellStyle);
            return clone;
        }

        /// <summary>
        /// 複製一筆預設的 表格表頭格格式
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitHeadCellStyle()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._HeadCellStyle);
            return clone;
        }

        /// <summary>
        /// 複製一筆預設的 表尾儲存格格式
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitFooterCellStyle()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._FooterCellStyle);
            return clone;
        }

        /// <summary>
        /// 複製一筆預設的 表身資料格格式(字串)
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitTableCellStyleForString()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._DataCellStyle[(int)DataCellDataTypeEnum.String]);
            return clone;
        }

        /// <summary>
        /// 複製一筆預設的 表身資料格格式(整數)
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitTableCellStyleForInt()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._DataCellStyle[(int)DataCellDataTypeEnum.NumberInt]);
            return clone;
        }


        /// <summary>
        /// 複製一筆預設的 表身資料格格式(有小數點的數字)
        /// </summary>
        /// <returns></returns>
        public XSSFCellStyle CloneInitTableCellStyleForNumberPoint()
        {
            XSSFCellStyle clone = (XSSFCellStyle)this._Workbook.CreateCellStyle();
            clone.CloneStyleFrom(this._DataCellStyle[(int)DataCellDataTypeEnum.NumberWithPoint]);
            return clone;
        }
        /// <summary>
        /// 建立一筆新的 字型格式
        /// </summary>
        /// <returns></returns>
        public XSSFFont CreateNewFont()
        {
            // 字型
            XSSFFont font = (XSSFFont)this._Workbook.CreateFont();
            font.FontName = "標楷體";
            font.FontHeightInPoints = 20;
            font.IsBold = false;

            return font;
        }
        /// <summary>
        ///  建立資料格
        /// </summary>
        /// <param name="row"></param>
        /// <param name="colIndex"></param>
        /// <param name="cellType"></param>
        /// <param name="cellStyle"></param>
        /// <returns></returns>
        private XSSFCell _CreateCell(XSSFRow row, int colIndex, CellType cellType, XSSFCellStyle cellStyle)
        {
            var cell = (XSSFCell)row.CreateCell(colIndex, cellType);
            cell.CellStyle = cellStyle;
            return cell;
        }
        #endregion

        #region 資料結構

        /// <summary>
        /// 報表各部分列舉
        /// </summary>
        private enum TableLayerEnum
        {
            /// <summary> 標題 </summary>
            Title = 1,
            /// <summary> 表頭 </summary>
            Head = 2,
            /// <summary> 表身 </summary>
            Data = 3,
            /// <summary> 表尾 </summary>
            Footer = 4
        }

        /// <summary>
        /// 資料格類型列舉
        /// </summary>
        private enum DataCellDataTypeEnum
        {
            /// <summary>數字整數</summary>
            NumberInt = 1,
            /// <summary>數字有小數</summary>
            NumberWithPoint = 2,
            /// <summary>字串</summary>
            String = 3,
        }
        /// <summary>
        /// 資料名稱資訊模型
        /// </summary>
        private class DataNameInfoModel
        {
            /// <summary>資料的屬性名稱</summary>
            public string DataName { get; set; }
            /// <summary>資料的屬性名稱</summary>
            public List<int> NeedMergeRowCount { get; set; }
            public List<MergeDataColumnModel> NeedMergeColumnCount { get; set; }
            /// <summary>該筆資料值如果是零的話，是否需要以空白顯示</summary>
            public bool ShowEmptyIfZero { get; set; }
            /// <summary>空白字串</summary>
            public string EmptyString { get; set; }
            /// <summary>本文字內容需套用的資料格格式(表身)(備註:目前只有string格式，如果有其他格式可能會出錯)</summary>
            public XSSFCellStyle ApplyCellStyleForColumn { get; set; }
            /// <summary>本文字內容需套用的對齊方式(表身)</summary>
            public HorizontalAlignment ApplyCellAlignmentForColumn { get; set; }
            /// <summary>資料列資料字體大小</summary>
            public short DataFontPoints { get; set; }
            /// <summary> 資料列資料小數點位數 </summary>
            public int? DataWithPoints { get; set; }
            /// <summary> 欄位裡是否string格式且要判別有數字 </summary>
            public bool IsStringWithNumber { get; set; }
            /// <summary>
            /// 建構子
            /// </summary>
            public DataNameInfoModel()
            {
                EmptyString = string.Empty;
                ApplyCellAlignmentForColumn = HorizontalAlignment.General;
            }
        }

        /// <summary>
        /// 資料名稱合併程序紀錄資訊模型
        /// </summary>
        private class DataNameMergeInfoModel
        {
            /// <summary>狀態：目前是否已執行合併</summary>
            public bool State { get; set; }
            /// <summary>指標：指定目前採用的合併資訊</summary>
            public int Pointer { get; set; }
            /// <summary>計數器：合併執行後，遞減已合併的欄位數量，直到Counter規0，重設狀態</summary>
            public int Counter { get; set; }
        }


        /// <summary>
        /// 表頭-表格關聯表模型(每一個bindModel 各代表一個欄位/資料格的資訊)
        /// 目前有支援 的資料類型為：(int, float, double, decimal, string)
        /// </summary>
        public class BindModel
        {
            /// <summary>本欄位名稱(表頭名稱)</summary>
            public string HeadName { get; set; }
            /// <summary>本文字內容需套用的資料格格式(表頭)</summary>
            public XSSFCellStyle ApplyCellStyle { get; set; }
            /// <summary>本文字內容需套用的資料格格式(表身)</summary>
            public XSSFCellStyle ApplyCellStyleForColumn { get; set; }
            /// <summary>本文字內容需套用的文字對齊方式(表身)</summary>
            public HorizontalAlignment ApplyCellAlignmentForColumn { get; set; }
            /// <summary>子欄位(子表頭資訊)</summary>
            public List<BindModel> SubDataBinding { get; set; }
            /// <summary>需輸出的input data model 的屬性名稱</summary>
            public List<string> DataName { get; set; }
            /// <summary>該欄需合併的列數</summary>
            public List<int> MergeDataRowCount { get; set; }
            /// <summary>
            /// 該列需合併的欄位資料
            /// </summary>
            public List<MergeDataColumnModel> MergeDataColumnCount { get; set; }
            /// <summary>該欄寬度</summary>
            public int? ColumnWidth { get; set; }
            /// <summary>為數字時，若為0是否輸出空白</summary>
            public bool ShowEmptyIfZero { get; set; }
            /// <summary>空白字串</summary>
            public string EmptyString { get; set; }
            /// <summary> 欄位裡是否string格式且要判別有數字 </summary>
            public bool IsStringWithNumber { get; set; }
            /// <summary>建構子</summary>
            public BindModel()
            {
                SubDataBinding = new List<BindModel>();
                DataName = new List<string>();
                MergeDataRowCount = new List<int>();
                EmptyString = string.Empty;
                MergeDataColumnCount = new List<MergeDataColumnModel>();
                ApplyCellAlignmentForColumn = HorizontalAlignment.General;
            }
            /// <summary>該欄所在列數(內部運算，請勿給值)</summary>
            public int RowNo { get; set; }
            /// <summary>該筆欄需要幾個欄位(內部運算，請勿給值)</summary>
            public int ColumsCount { get; set; }
            /// <summary>該欄之下有幾列(內部運算，請勿給值)</summary>
            public int NextRowCount { get; set; }
            /// <summary>與本欄有相關的總列數(內部運算，請勿給值)</summary>
            public int MaxRowCount { get; set; }
            /// <summary>資料列表頭字體大小</summary>
            public short HeadFontPoints { get; set; }
            /// <summary>資料列資料字體大小</summary>
            public short DataFontPoints { get; set; }
            /// <summary> 資料列資料小數點位數 </summary>
            public int? DataWithPoints { get; set; }
        }
        /// <summary>
        /// 欄位合併模型
        /// </summary>
        public class MergeDataColumnModel
        {
            /// <summary>
            /// 要合併之開始欄-索引
            /// </summary>
            public int ColumnIdx { get; set; }
            /// <summary>
            /// 要合併之開始列-索引
            /// </summary>
            public int RowIdx { get; set; }
            /// <summary>
            /// 要合併之欄位數
            /// </summary>
            public int MergeColumnCount { get; set; }
            /// <summary>
            /// 合併完後的資料格對齊格式
            /// </summary>
            public HorizontalAlignment MergeCellAlignType { get; set; }
            /// <summary> 合併完後的資料格字體大小 </summary>
            public short MergeCellPoints { get; set; }
        }
        /// <summary>
        /// 特定一列的 表頭/表尾 資料模型
        /// </summary>
        public class TitleModel
        {
            /// <summary>本列 需顯示的文字內容</summary>
            public List<TitleContentModel> Content { get; set; }
            /// <summary>本列 高度(-1：表預設值)</summary>
            public int RowHeightInPoint { get; set; }

            /// <summary>
            /// 建構子
            /// </summary>
            public TitleModel()
            {
                RowHeightInPoint = -1;
            }
        }
        /// <summary>
        /// 表頭/表尾 文字內容模型
        /// </summary>
        public class TitleContentModel
        {
            /// <summary>顯示文字</summary>
            public string Text { get; set; }
            /// <summary>需合併列數(不含本列)</summary>
            public int? MergeRowCount { get; set; }
            /// <summary>需合併欄數(不含本欄)</summary>
            public int? MergeColumnCount { get; set; }
            /// <summary>本文字內容需套用的資料格格式</summary>
            public XSSFCellStyle ApplyCellStyle { get; set; }

        }

        /// <summary>
        /// 分頁資訊模型
        /// </summary>
        /// <typeparam name="DataModel"></typeparam>
        public class SheetInfoModel<DataModel> where DataModel : IReport
        {
            /// <summary> 分頁名稱 </summary>
            public string Name { get; set; }
            /// <summary> 列印時紙張大小「-1：預設，不調整」、「 8：A3」、「 8：A3」、「 9：A4」、「1：Letter」</summary>
            public short PrintPageSize { get; set; }
            /// <summary> 列印時/指定直式或橫式「true：橫式」、「false：直式」，預設 true</summary>
            public bool PrintLandScape { get; set; }
            /// <summary>標題資料集</summary>
            protected List<TitleModel> _DataTitle { get; set; }
            /// <summary>表格綁定資訊</summary>
            private List<BindModel> _DataBinding { get; set; }
            /// <summary>表尾列資料集</summary>
            private List<TitleModel> _DataFooter { get; set; }
            /// <summary>表格資料列</summary>
            private List<DataModel> _DataList { get; set; }

            /// <summary>
            /// 建構子
            /// </summary>
            public SheetInfoModel()
            {
                // 初始化
                this._DataBinding = new List<BindModel>();
                this._DataTitle = new List<TitleModel>();
                this._DataFooter = new List<TitleModel>();
                this.PrintPageSize = -1;
                this.PrintLandScape = true;
            }

            #region 報表資料設定
            /// <summary>
            /// 設定 表頭-表身 資料關聯
            /// </summary>
            /// <param name="BindModel"></param>
            public void SetBindModel(List<BindModel> BindModel)
            {
                this._DataBinding = BindModel;


            }
            /// <summary>
            /// 設定 表尾列資料
            /// </summary>
            /// <param name="footerList"></param>
            public void SetFooterModel(List<TitleModel> footerList)
            {
                this._DataFooter = footerList;
            }
            /// <summary>
            /// 設定 表格資料
            /// </summary>
            /// <param name="dataList"></param>
            public void SetDate(List<DataModel> dataList)
            {
                this._DataList = dataList;

            }
            /// <summary>
            /// 設定 標題資料
            /// </summary>
            /// <param name="titleList"></param>
            public void SetTitleModel(List<TitleModel> titleList)
            {
                this._DataTitle = titleList;
            }

            #endregion

            #region 取得報表資料
            /// <summary>
            /// 取得 標題資料
            /// </summary>
            /// <returns></returns>
            public List<TitleModel> GetTitleList()
            {
                return new List<TitleModel>(this._DataTitle);
            }
            /// <summary>
            /// 取得 表頭-表身的關聯表資訊
            /// </summary>
            /// <returns></returns>
            public List<BindModel> GetDataBindingList()
            {
                return new List<BindModel>(this._DataBinding);
            }
            /// <summary>
            /// 取得 表尾資料
            /// </summary>
            /// <returns></returns>
            public List<TitleModel> GetDataFooterList()
            {
                return new List<TitleModel>(this._DataFooter);
            }
            /// <summary>
            /// 取得 表身資料
            /// </summary>
            /// <returns></returns>
            public List<DataModel> GetDataList()
            {
                return new List<DataModel>(this._DataList);
            }
            #endregion
        }
        #endregion

        /// <summary>
        /// 取得MIME TYPE
        /// </summary>
        /// <returns></returns>
        public virtual string GetMimeType()
        {
            return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        }
        /// <summary>
        /// 取得副檔名
        /// </summary>
        /// <returns></returns>
        public virtual string GetExtension()
        {
            return "xlsx";
        }

    }

    /// <summary>
    /// 報表匯出工具 資料介面
    /// </summary>
    public interface IReport
    {

    }
    #endregion
}