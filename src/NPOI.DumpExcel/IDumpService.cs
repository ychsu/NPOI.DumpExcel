using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;

namespace NPOI.DumpExcel
{
    public interface IDumpService<T>
    {
        /// <summary>
        /// 匯出 enumerable 到 excel workbook
        /// </summary>
        /// <param name="enumerable"></param>
        /// <returns></returns>
        IWorkbook DumpWorkbook(IEnumerable<T> enumerable);
    }
}
