namespace Reporter.Models
{
    /// <summary>
    /// 分頁產生器
    /// </summary>
    public abstract class ReportSheetCreater
    {
        private delegate bool EmptyCheckDelegation(object value);

        /// <summary>分頁物件</summary>
        protected NPOIExportTool.SheetInfoModel<IReport> Sheet = new NPOIExportTool.SheetInfoModel<IReport>();
        /// <summary>資料表總欄數</summary>
        public int TotalColumnCount { get; set; }
        /// <summary>資料表總列數</summary>
        public int TotalDataRowCount { get; set; }
        /// <summary>該分頁的匯出工具</summary>
        protected NPOIExportTool Exporter { get; set; }
        /// <summary>該分頁的表格資料</summary>
        protected List<IReport> DataList { get; set; }

        /// <summary>設定 表頭-表身 資料繫節</summary>
        protected abstract List<NPOIExportTool.BindModel> SetupTableBinding();
        /// <summary>設定 標題資料 </summary>
        protected abstract List<NPOIExportTool.TitleModel> SetupTitle();
        /// <summary>設定 表尾資料 </summary>
        protected abstract List<NPOIExportTool.TitleModel> SetupFooter();
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="exporter"></param>
        /// <param name="sheetName"></param>
        /// <param name="data"></param>
        public ReportSheetCreater(NPOIExportTool exporter, string sheetName, List<IReport> data)
        {

            this.DataList = data;
            this.Exporter = exporter;
            this.Sheet.Name = sheetName;
        }

        /// <summary>
        /// 綁定資料表
        /// </summary>
        protected void BindAllData()
        {
            this.TotalDataRowCount = this.DataList.Count();
            var bindingModel = SetupTableBinding();
            foreach (var item in bindingModel)
            {
                this._BuildLayerInfo(item, 1);
            }
            this.TotalColumnCount = bindingModel.Sum(x => x.ColumsCount);

            this.Sheet.SetDate(this.DataList);
            this.Sheet.SetBindModel(bindingModel);
            this.Sheet.SetTitleModel(SetupTitle());
            this.Sheet.SetFooterModel(SetupFooter());
            this.Exporter.AddSheet(this.Sheet);
        }

        /// <summary>
        /// 欄位是否全為空值檢查
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns>Dictionary[欄位名稱] = 是否全空</returns>
        protected Dictionary<string, bool> _EmptyColumnDic<T>(List<IReport> data) where T : new()
        {

            return _EmptyColumnDic<T>(data, new List<string>());
        }

        /// <summary>
        /// 欄位是否全為空值檢查
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="ignoreColumnNames"></param>
        /// <returns>Dictionary[欄位名稱] = 是否全空</returns>
        protected Dictionary<string, bool> _EmptyColumnDic<T>(List<IReport> data, List<string> ignoreColumnNames) where T : new()
        {

            var stringTag = typeof(string).Name;
            var intTag = typeof(int).Name;
            var decimalTag = typeof(decimal).Name;
            var emptyDetectorDic = new Dictionary<string, EmptyCheckDelegation> {
                { stringTag, new EmptyCheckDelegation(_StringHandler)},
                { decimalTag, new EmptyCheckDelegation(_DecimalHandler)},
                { intTag, new EmptyCheckDelegation(_IntHandler)},
            };
            var properties = (new T()).GetType().GetProperties();
            if (!data.Any())
                return properties.ToDictionary(x => x.Name, x => true);


            var dataIsNotEmptyDic = properties.ToDictionary(x => x.Name, x => false);
            ignoreColumnNames = ignoreColumnNames ?? new List<string>();
            foreach (var item in data)
            {
                foreach (var property in properties)
                {
                    var existThisProperty = dataIsNotEmptyDic.ContainsKey(property.Name);
                    if (ignoreColumnNames.Contains(property.Name) && existThisProperty)
                    {
                        dataIsNotEmptyDic[property.Name] = true;
                        continue;
                    }


                    if (!existThisProperty || dataIsNotEmptyDic[property.Name])
                        continue;

                    var propertyTypeName = property.PropertyType.Name;
                    if (emptyDetectorDic.ContainsKey(propertyTypeName))
                    {
                        var value = property.GetValue(item);
                        var detector = emptyDetectorDic[propertyTypeName];
                        var isEmpty = detector(value);
                        dataIsNotEmptyDic[property.Name] = !isEmpty;
                    }
                    else
                    {
                        dataIsNotEmptyDic[property.Name] = true;
                    }
                }
            }
            return dataIsNotEmptyDic;
        }

        #region 其他私有運算

        /// <summary>
        /// 檢查文字是否為空
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _StringHandler(object value)
        {
            return string.IsNullOrWhiteSpace((string)value);
        }


        /// <summary>
        /// 檢查數值是否為0
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _DecimalHandler(object value)
        {
            return (decimal)value == 0;
        }

        /// <summary>
        /// 檢查數值是否為0
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool _IntHandler(object value)
        {
            return (int)value == 0;
        }

        /// <summary>
        /// 建立分層資訊
        /// </summary>
        /// <param name="data">需分層的關聯表</param>
        /// <param name="nowRow">目前所在列數</param>
        private void _BuildLayerInfo(NPOIExportTool.BindModel data, int nowRow)
        {
            // 底下無其他子表頭時:設定本欄位的分層資訊
            if (!data.SubDataBinding.Any())
            {
                data.RowNo = nowRow;
                data.ColumsCount = data.DataName.Count();
                data.NextRowCount = 0;
                data.MaxRowCount = nowRow;
            }
            else
            {
                // 為其子欄位建立分層資訊
                foreach (var item in data.SubDataBinding)
                {
                    _BuildLayerInfo(item, nowRow + 1);
                }


                data.RowNo = nowRow;
                // 本欄位的欄位需求量 = 其下一層的所需欄位總和
                data.ColumsCount = data.SubDataBinding.Sum(x => x.ColumsCount);
                // 與本欄位關聯的總列數
                data.MaxRowCount = data.SubDataBinding.Max(x => x.MaxRowCount);
                // 本欄位下還存有的總列數
                data.NextRowCount = data.MaxRowCount - nowRow;

            }
        }
        #endregion

        /// <summary>
        /// 取得分頁
        /// </summary>
        /// <returns></returns>
        public NPOIExportTool.SheetInfoModel<IReport> GetSheet()
        {
            return this.Sheet;
        }

        /// <summary> 是否使用ODS匯出 </summary>
        /// <returns></returns>
        public bool GetIsOds()
        {
            return Exporter.GetType().ToString().EndsWith("Ods");
        }
        /// <summary> 是否使用PDF匯出 </summary>
        /// <returns></returns>
        public bool GetIsPdf()
        {
            return Exporter.GetType().ToString().EndsWith("Pdf");
        }


        /// <summary>
        ///  設定分頁資訊(預設為A4、直式)
        /// </summary>
        /// <param name="PrintLandScape">列印橫式(true)/直式(false)</param>
        /// <param name="PrintPageSize">紙張大小 => (8:A3)、(9:A4)</param>
        public void SetSheet(bool? PrintLandScape = null, short? PrintPageSize = null)
        {
            bool _printLandScape = PrintLandScape ?? false;
            short _printPageSize = PrintPageSize ?? 9;
            var sheetEntity = GetSheet();
            sheetEntity.PrintLandScape = _printLandScape;
            sheetEntity.PrintPageSize = _printPageSize;
        }

        /// <summary>
        /// C系列報表字體大小調整
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public ReportParametersModel SetCReportFont(ReportParametersModel parameter)
        {
            parameter.titleFontPoints = 18;
            parameter.DataFontPoints = 14;
            parameter.printerFontPoints = 14;

            var headfont = parameter.npoiExportTool.CreateNewFont();
            headfont.FontHeightInPoints = 14;
            headfont.IsBold = true;
            parameter.headCellStyle.SetFont(headfont);

            var footfont = parameter.npoiExportTool.CreateNewFont();
            footfont.FontHeightInPoints = 14;
            parameter.footCellStyle.SetFont(footfont);

            return parameter;
        }


        /// <summary>
        /// 產生空白鍵字串
        /// </summary>
        /// <param name="count"></param>
        /// <returns></returns>
        public string CreateSpaces(int count)
        {
            var aSpace = " ";
            var result = "";
            while (count > 0)
            {
                result += aSpace;
                count--;
            }

            return result;

        }

        #region A3、A4固定格式模型

        /// <summary>
        /// 報表參數模型(基底)
        /// </summary>
        public abstract class ReportParametersModel
        {
            /// <summary> npoi工具</summary>
            public NPOIExportTool npoiExportTool { get; set; }
            /// <summary>列印格式(true : 橫式 ； false : 直式)</summary>
            public bool printLandScape { get; set; }
            /// <summary>列印紙張大小(-1 : 不調整 ； 8 : A3； 9 : A4)</summary>
            public short printPageSize { get; set; }
            /// <summary>標頭列(是否粗體)</summary>
            public bool titleBold { get; set; }
            /// <summary>標頭列(對齊方式)</summary>
            public HorizontalAlignment titleAlignment { get; set; }
            /// <summary>標頭列(字體大小)</summary>
            public short titleFontPoints { get; set; }
            /// <summary>列印日期列(是否粗體)</summary>
            public bool printerBold { get; set; }
            /// <summary>列印日期列(對齊方式)</summary>
            public HorizontalAlignment printerAlignment { get; set; }
            /// <summary>列印日期列(字體大小)</summary>
            public short printerFontPoints { get; set; }
            /// <summary>表頭列格式</summary>
            public XSSFCellStyle headCellStyle { get; set; }
            /// <summary>資料列(字體大小)</summary>
            public short DataFontPoints { get; set; }
            /// <summary>數字欄位是否要靠右對齊</summary>
            public bool isAlignRightForNum { get; set; }
            /// <summary>結尾列格式</summary>
            public XSSFCellStyle footCellStyle { get; set; }

            /// <summary>其他標題列設定</summary>
            public Dictionary<string, OtherTitleContentCellStyleModel> OtherCellStyleDic { get; set; }
            /// <summary>(標題列)全員當年KEY值</summary>
            public string allYearTitleKey { get { return "TITLE-ALLYEAR"; } }
            /// <summary>(標題列)全員當月KEY值</summary>
            public string allMonthTitleKey { get { return "TITLE-ALLMONTH"; } }
            /// <summary>(標題列)個人當月KEY值</summary>
            public string personalMonthTitleKey { get { return "TITLE_PERSONALMONTH"; } }

            /// <summary>(結尾列)主辦、主管、秘書、單位首長KEY值</summary>
            public string hostFootKey { get { return "FOOT-HOST"; } }
            /// <summary>(結尾列)主辦、主任、負責人KEY值</summary>
            public string hostForeignerFootKey { get { return "FOOT-FEIGNER-HOST"; } }
            /// <summary> (結尾列)主辦、主任、秘書、單位首長KEY值(FOR-C04[10個字元]) </summary>
            public string hostFootKeyForC04WithTenSpace { get { return "FOOT-HOST-TEN-SPACE"; } }
            /// <summary>(結尾列)主辦、主任、秘書、單位首長KEY值(FOR-C03[10個字元])</summary>
            public string hostFootKeyForC03WithTwentySpace { get { return "FOOT-HOST-TWENTY-SPACE"; } }
            /// <summary> (結尾列)備註：產業專業團包含茶業團、畜牧飼育團、菇蕈團、果樹團及設施協作團KEY值(FOR-B02) </summary>
            public string proGroupTipRemarkKey { get { return "FOOT-PRO-GROUP-TEA-REMARK"; } }
            /// <summary> (結尾列)主辦、主任、秘書、單位首長KEY值(FOR-A02 ODS、PDF[5個字元]) </summary>
            public string hoosFootKeyForODSPDF { get { return "FOOT-HOST-FIVE-SPACE-ODSPDF"; } }
            /// <summary>(結尾列)執行率說明備註</summary>
            public string executiveRateRemarkFootKey { get { return "FOOT-EXE-RATE-REMAKR-KEY"; } }
            /// <summary>(標題列、結尾列) 可接受之KEY值清單</summary>
            public List<string> canAcceptTitleKeyList { get; set; }
            /// <summary>簽名者之間相隔全形空白鍵數量(fullSpaceBetweenSignersDic[Key] = fullSpaceCount)</summary>
            public Dictionary<string, int> fullSpaceBetweenSignersDic { get; set; }
            /// <summary>
            /// 基底建構子
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            /// <param name="PrintLandScape">橫/直式</param>
            /// <param name="PrintPageSize">報表列印格式(A3、A4)</param>
            public ReportParametersModel(NPOIExportTool Tool, bool PrintLandScape, short PrintPageSize)
            {
                this.npoiExportTool = Tool;
                this.printLandScape = PrintLandScape;
                this.printPageSize = PrintPageSize;

                this.OtherCellStyleDic = new Dictionary<string, OtherTitleContentCellStyleModel>();
                this.canAcceptTitleKeyList = new List<string>();
                this.fullSpaceBetweenSignersDic = new Dictionary<string, int>();
            }
            /// <summary>
            /// 基底抽象方法(設定參數)
            /// </summary>
            public abstract void _SetParameters();
            /// <summary>
            /// 基底抽象方法(取得"全員當月"格式)
            /// </summary>
            /// <returns></returns>
            public abstract OtherTitleContentCellStyleModel _GetAllMonthCellStyle();
            /// <summary>
            /// 基底抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public abstract OtherTitleContentCellStyleModel _GetAllYearCellStyle();
            /// <summary>
            /// 基底抽象方法(取得"主辦、主任、秘書、單位首長"格式)
            /// </summary>
            /// <returns></returns>
            public abstract OtherTitleContentCellStyleModel _GetHostCellStyle();
            /// <summary>
            /// 基底抽象方法(取得"主辦、主管、負責人"格式)
            /// </summary>
            /// <returns></returns>
            public abstract OtherTitleContentCellStyleModel _GetHostForeignerCellStyle();
            /// <summary>
            /// 基底抽象方法(取得"個人當月"格式)
            /// </summary>
            /// <returns></returns>
            public abstract OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle();

            /// <summary>
            /// 取得執行率說明備註格式
            /// </summary>
            /// <returns></returns>
            public virtual OtherTitleContentCellStyleModel _GetExecutiveRateRemarkCellStyle()
            {
                return new OtherTitleContentCellStyleModel
                {
                    Name = "執行率% = 實際派工人天/應上工天數*100%(應上工天數不包含農務人員當月請假、農業專業訓練)",
                    IsTitleArea = false,
                    Style = footCellStyle,
                    IsMergeAllColumn = true,
                };
            }


            /// <summary>
            /// 一般簽名者清單
            /// </summary>
            public List<string> GeneralSigners = new List<string> { "主辦：", "主任：", "秘書：", "單位首長：" };

            /// <summary>
            /// 外展簽名者清單
            /// </summary>
            public List<string> ForeignerSigners = new List<string> { "主辦：", "主任：", "負責人：" };

            /// <summary>
            ///  產生簽名者頁尾
            /// </summary>
            /// <param name="signers">簽名者清單</param>
            /// <param name="fullSpaceCounts">簽名者間相隔全形空白鍵數量</param>
            /// <returns></returns>
            public string CreateSingerFooter(List<string> signers, int fullSpaceCounts)
            {
                var aFullSpace = "　";
                var spacesString = "";
                while (fullSpaceCounts > 0)
                {

                    spacesString += aFullSpace;
                    fullSpaceCounts--;
                }
                return string.Join(spacesString, signers);
            }


        }
        /// <summary>
        /// 橫向、A3格式 報表參數模型 (Title:38、22pt，Content:24pt， Footer:26pt)
        /// </summary>
        public class A3HorizontalModel : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(橫向、A3格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A3HorizontalModel(NPOIExportTool Tool) : base(Tool, true, 8)
            {
                _SetParameters();
            }

            /// <summary>
            /// 實作抽象方法(取得"全員當月"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 24;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 24;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 25個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 25);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 37個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {

                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 37);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"個人當月"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var personalMonthFont = base.npoiExportTool.CreateNewFont();
                var personalMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                personalMonthFont.FontHeightInPoints = 24;
                personalMonthFont.IsBold = false;
                personalMonthCellStyle.SetFont(personalMonthFont);
                personalMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = personalMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }


            /// <summary>
            /// 設定A3格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 38;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 22;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 24;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 24;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 26;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(hostForeignerFootKey, _GetHostForeignerCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
                base.OtherCellStyleDic.Add(executiveRateRemarkFootKey, _GetExecutiveRateRemarkCellStyle());
            }

        }

        /// <summary>
        /// 橫向、A3格式 報表參數模型 (Title:38、22pt，Content:24pt， Footer:26pt)
        /// </summary>
        public class A3HorizontalShortHostModel : A3HorizontalModel
        {
            /// <summary>
            /// 報表參數模型(橫向、A3格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A3HorizontalShortHostModel(NPOIExportTool Tool) : base(Tool)
            {

            }


            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 15個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {

                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 15);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 22個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 22);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

        }

        /// <summary>
        /// 橫向、A4格式 報表參數模型 (Title:26、18pt，Content:18pt， Footer:18pt)
        /// </summary>
        public class A4HorizontalModel : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(橫向、A4格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A4HorizontalModel(NPOIExportTool Tool) : base(Tool, true, 9)
            {
                _SetParameters();
            }
            /// <summary>
            /// 實作抽象方法(取得全員當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 15個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 15);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作方法(取得"主辦、主任、秘書、單位首長" 10個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public OtherTitleContentCellStyleModel _GetHostCellStyleWithTenSpace()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKeyForC04WithTenSpace, 10);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKeyForC04WithTenSpace];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 17個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 17);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得個人當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作方法(取得"備註：產業專業團包含茶業團、畜牧飼育團、菇蕈團、果樹團及設施協作團" [B02] )
            /// </summary>
            /// <returns></returns>
            public OtherTitleContentCellStyleModel _GetProGroupTipRemarkCellStyle()
            {
                return new OtherTitleContentCellStyleModel
                {
                    Name = "備註：產業專業團包含茶業團、畜牧飼育團、菇蕈團、果樹團及設施協作團",
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 設定A4格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                // todo 調整規格
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 26;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 18;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 18;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 18;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 18;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(hostFootKeyForC04WithTenSpace, _GetHostCellStyleWithTenSpace());
                base.OtherCellStyleDic.Add(hostForeignerFootKey, _GetHostForeignerCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
                base.OtherCellStyleDic.Add(proGroupTipRemarkKey, _GetProGroupTipRemarkCellStyle());
                base.OtherCellStyleDic.Add(executiveRateRemarkFootKey, _GetExecutiveRateRemarkCellStyle());
            }

        }
        /// <summary>
        /// 直向、A4格式 報表參數模型 (Title:22、18pt，Content:18pt， Footer:18pt)
        /// </summary>
        public class A4StraightModel : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(直向、A4格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A4StraightModel(NPOIExportTool Tool) : base(Tool, false, 9)
            {
                _SetParameters();
            }
            /// <summary>
            /// 實作抽象方法(取得全員當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 7個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 7);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 10個空白全形格式)
            /// </summary>
            /// <returns></returns>
            public OtherTitleContentCellStyleModel _GetHostCellStyleWithTwentySpace()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKeyForC04WithTenSpace, 10);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKeyForC04WithTenSpace];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 15個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 15);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作方法(取得"主辦、主任、秘書、單位首長" 5個全形空白格式)(for pdf、ods)
            /// </summary>
            /// <returns></returns>
            public OtherTitleContentCellStyleModel _GetHostCellStyleWithFiveSpace()
            {


                base.fullSpaceBetweenSignersDic.Add(hoosFootKeyForODSPDF, 5);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hoosFootKeyForODSPDF];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得個人當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 設定A4格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                // todo 調整規格
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 22;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 18;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 18;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 18;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 18;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(hostFootKeyForC04WithTenSpace, _GetHostCellStyleWithTwentySpace());
                base.OtherCellStyleDic.Add(hoosFootKeyForODSPDF, _GetHostCellStyleWithFiveSpace());
                base.OtherCellStyleDic.Add(hostForeignerFootKey, _GetHostForeignerCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
            }

        }

        /// <summary>
        /// 橫向、A4格式 報表參數模型 (Title:44、26pt，Content:28pt， Footer:30pt)
        /// </summary>
        public class A4HorizontalModelForA03FarmMechanical : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(橫向、A4格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A4HorizontalModelForA03FarmMechanical(NPOIExportTool Tool) : base(Tool, true, 9)
            {
                _SetParameters();
            }
            /// <summary>
            /// 實作抽象方法(取得全員當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長" 17個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {

                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 17);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 27個全形空白格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 27);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得個人當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 設定A4格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                // todo 調整規格
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 44;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 26;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 28;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 28;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 30;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
            }

        }

        /// <summary>
        /// 直向、A4格式 報表參數模型 (Title:20、18pt，Content:18pt， Footer:18pt)
        /// </summary>
        public class A4StraightSmallReportParametersModel : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(直向、A4格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A4StraightSmallReportParametersModel(NPOIExportTool Tool) : base(Tool, false, 9)
            {
                _SetParameters();
            }
            /// <summary>
            /// 實作抽象方法(取得全員當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長"格式 8個全形空白鍵)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 8);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人" 12個全形空白鍵)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 12);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);

                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得個人當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 設定A4格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 20;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 18;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 18;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 18;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 18;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(hostForeignerFootKey, _GetHostForeignerCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
            }

        }

        /// <summary>
        /// 直向、A4格式 報表參數模型 (Title:32、18pt，Content:20pt， Footer:20pt)
        /// </summary>
        public class A4StraightModelForC03 : ReportParametersModel
        {
            /// <summary>
            /// 報表參數模型(直向、A4格式)
            /// </summary>
            /// <param name="Tool">NPOI匯出工具</param>
            public A4StraightModelForC03(NPOIExportTool Tool) : base(Tool, false, 9)
            {
                _SetParameters();
            }
            /// <summary>
            /// 實作抽象方法(取得全員當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"全員當年"格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetAllYearCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "全員當年",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長"格式 15個全形空白鍵)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKey, 15);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKey];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }
            /// <summary>
            /// 實作抽象方法(取得"主辦、主任、秘書、單位首長"格式)(10個全形空白)
            /// </summary>
            /// <returns></returns>
            public OtherTitleContentCellStyleModel _GetHostCellStyleWithTwentySpace()
            {
                base.fullSpaceBetweenSignersDic.Add(hostFootKeyForC03WithTwentySpace, 10);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostFootKeyForC03WithTwentySpace];
                var signerContent = CreateSingerFooter(GeneralSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得"主辦、主管、負責人") (20個全形空白鍵)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetHostForeignerCellStyle()
            {
                base.fullSpaceBetweenSignersDic.Add(hostForeignerFootKey, 20);
                var fullSpaceCount = base.fullSpaceBetweenSignersDic[hostForeignerFootKey];
                var signerContent = CreateSingerFooter(ForeignerSigners, fullSpaceCount);
                return new OtherTitleContentCellStyleModel
                {
                    Name = signerContent,
                    IsTitleArea = false,
                    Style = base.footCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 實作抽象方法(取得個人當月格式)
            /// </summary>
            /// <returns></returns>
            public override OtherTitleContentCellStyleModel _GetPersonalMonthCellStyle()
            {
                var allMonthFont = base.npoiExportTool.CreateNewFont();
                var allMonthCellStyle = base.npoiExportTool.CloneInitCellStyle();
                allMonthFont.FontHeightInPoints = 16;
                allMonthFont.IsBold = false;
                allMonthCellStyle.SetFont(allMonthFont);
                allMonthCellStyle.Alignment = HorizontalAlignment.Left;
                return new OtherTitleContentCellStyleModel
                {
                    Name = "個人當月",
                    IsTitleArea = true,
                    Style = allMonthCellStyle,
                    IsMergeAllColumn = true,
                };
            }

            /// <summary>
            /// 設定A4格式報表參數
            /// </summary>
            public override void _SetParameters()
            {
                // todo 調整規格
                base.titleBold = true;
                base.titleAlignment = HorizontalAlignment.Center;
                base.titleFontPoints = 32;

                base.printerBold = false;
                base.printerAlignment = HorizontalAlignment.Right;
                base.printerFontPoints = 18;

                var cloneHeadCellStyle = base.npoiExportTool.CloneInitHeadCellStyle(); //粗體 標楷體 置中 16pt
                var headfont = base.npoiExportTool.CreateNewFont();
                headfont.IsBold = true;
                headfont.FontHeightInPoints = 20;
                cloneHeadCellStyle.SetFont(headfont);
                base.headCellStyle = cloneHeadCellStyle;

                base.DataFontPoints = 20;
                base.isAlignRightForNum = true;

                var cloneFootCellStyle = base.npoiExportTool.CloneInitFooterCellStyle(); // 標楷體 靠左 16pt
                var footfont = base.npoiExportTool.CreateNewFont();
                footfont.FontHeightInPoints = 20;
                cloneFootCellStyle.SetFont(footfont);
                base.footCellStyle = cloneFootCellStyle;

                base.OtherCellStyleDic.Add(allMonthTitleKey, _GetAllMonthCellStyle());
                base.OtherCellStyleDic.Add(allYearTitleKey, _GetAllYearCellStyle());
                base.OtherCellStyleDic.Add(hostFootKey, _GetHostCellStyle());
                base.OtherCellStyleDic.Add(hostFootKeyForC03WithTwentySpace, _GetHostCellStyleWithTwentySpace());
                base.OtherCellStyleDic.Add(hostForeignerFootKey, _GetHostForeignerCellStyle());
                base.OtherCellStyleDic.Add(personalMonthTitleKey, _GetPersonalMonthCellStyle());
            }
        }

        /// <summary>
        /// 其他標題列模型
        /// </summary>
        public class OtherTitleContentCellStyleModel
        {
            /// <summary>名稱</summary>
            public string Name { get; set; }
            /// <summary>格式</summary>
            public XSSFCellStyle Style { get; set; }
            /// <summary>是否要放在頁首標題列(否則頁尾)</summary>
            public bool IsTitleArea { get; set; }
            /// <summary>是否要合併整欄</summary>
            public bool IsMergeAllColumn { get; set; }
            /// <summary>合併欄數(不含本欄)</summary>
            public int MergeColumnCount { get; set; }
            /// <summary>合併列數(不含本列)</summary>
            public int MergeRowCount { get; set; }

        }

        #endregion





    }
}