using NPOI.DumpExcel.Structs;
using NPOI.SS.UserModel;
using System;

namespace NPOI.DumpExcel.Attributes
{
    /// <summary>
    /// excel 欄位設定
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// excel欄位設定
        /// </summary>
        /// <param name="Order">excel 欄位 排序</param>
        /// <param name="Format">cell format</param>
        /// <param name="Width">column width</param>
        /// <param name="VerticalAlign">垂直位置</param>
        /// <param name="Align">水平位置</param>
        public ExcelColumnAttribute(int Order = -1,
            string Format = "",
            int Width = -1,
            VerticalAlignment VerticalAlign = VerticalAlignment.Center,
            HorizontalAlignment Align = HorizontalAlignment.Center)
        {
            this.CellFormatter = new CellFormatter
            {
                Alignment = Align,
                Format = Format,
                VerticalAlign = VerticalAlign,
            };
            this.Width = Width;
            this.Order = Order < 0 ? int.MaxValue : Order;
        }

        /// <summary>
        /// cell formatter
        /// </summary>
        internal CellFormatter CellFormatter { get; set; }

        /// <summary>
        /// column width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// column sort threshold
        /// </summary>
        public int Order { get; set; }
    }
}
