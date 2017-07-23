using NPOI.DumpExcel.Attributes;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.DumpExcel.Test.Models
{

    [Sheet(Name = "Foo Sheet")]
    public class Foo
    {
        [ExcelColumn(Format: "#.00", Width: 20, Align: HorizontalAlignment.Right)]
        public int SerId { get; set; }
        [ExcelColumn]
        public string Name { get; set; }
        [ExcelColumn]
        public Enum1 Enum0 { get; set; }
        [ExcelColumn]
        public Enum1 Enum1 { get; set; }
        [ExcelColumn(Format: "yyyy-MM-dd")]
        public DateTime DT { get; set; }
    }

    public enum Enum1
    {
        AAAAAAAAAAAAAAAAAAAAAA,
        BBBBBBBBBBBBBBBBBBBBBB
    }
}