using NPOI.SS.UserModel;

namespace NPOI.DumpExcel.Structs
{
    internal struct CellFormatter
    {
        /// <summary>
        /// 顯示格式
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// horizontal alignment
        /// </summary>
        public HorizontalAlignment Alignment { get; set; }


        /// <summary>
        /// vertical align
        /// </summary>
        public VerticalAlignment VerticalAlign { get; set; }
    }
}
