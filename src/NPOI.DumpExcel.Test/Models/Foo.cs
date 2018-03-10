using NPOI.DumpExcel.Attributes;
using NPOI.SS.UserModel;
using System;
using System.ComponentModel.DataAnnotations;

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
        [ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
        public DateTimeOffset? DT { get; set; }

        [ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
        public DateTime DT2 { get; set; } = DateTime.Now;
    }

    public enum Enum1
    {
        [Display(Name = "我是A")]
        AAAAAAAAAAAAAAAAAAAAAA,
        BBBBBBBBBBBBBBBBBBBBBB
    }
}