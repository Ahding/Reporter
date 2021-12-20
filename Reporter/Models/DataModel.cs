using Reporter.Attributes;
using Reporter.Tools;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reporter.Models
{
    /// <summary>
    /// 匯出模型範例
    /// </summary>
    public class DataModel : IReport
    {
        /// <summary>單位</summary>
        [Description("單位")]
        public string Organization { get; set; }

        /// <summary>ID</summary>
        [Description("ID")]
        public string ID { get; set; }


        /// <summary>姓名</summary>
        [Description("姓名")]
        public string Name { get; set; }

        /// <summary>備註</summary>
        [ColumnWidth(ColumnWidth = 100)]
        [Description("備註")]
        public string Remark { get; set; }
        /// <summary>備註2</summary>
        //[ColumnWidth(ColumnWidth = 100)]
        //[Description("備註2")]
        //public string Remark2 { get; set; }
    }
}
