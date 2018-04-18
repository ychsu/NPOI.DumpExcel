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
        public CaseType? Enum1 { get; set; }
        [ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
        public DateTimeOffset? DT { get; set; }

        [ExcelColumn(Width: 20)]
		public bool Bool { get; set; }

		[ExcelColumn(Format: "yyyy-MM-dd", Width: 20)]
        public DateTime DT2 { get; set; } = DateTime.Now;
    }

    public enum Enum1
    {
        [Display(Name = "我是A")]
        AAAAAAAAAAAAAAAAAAAAAA,
        BBBBBBBBBBBBBBBBBBBBBB
	}
	public enum CaseType
	{
		[Display(Name = "服務")]
		Service,
		[Display(Name = "法規")]
		Regulation,
		[Display(Name = "測試")]
		Test,
		[Display(Name = "報告")]
		Report,
		[Display(Name = "其他(備註)")]
		Other
	}
}