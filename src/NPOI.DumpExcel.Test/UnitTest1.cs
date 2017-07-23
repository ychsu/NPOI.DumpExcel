using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using NPOI.DumpExcel.Test.Models;
using System.IO;

namespace NPOI.DumpExcel.Test
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void DumpEnumerableToExcel()
        {
            var enumerable = Enumerable.Range(1, 10)
                .Select(p => new Foo
                {
                    DT = DateTime.Now.AddDays(p),
                    Enum0 = Enum1.AAAAAAAAAAAAAAAAAAAAAA,
                    Enum1 = Enum1.BBBBBBBBBBBBBBBBBBBBBB,
                    Name = $"Foo{p}",
                    SerId = p
                });

            var workbook = enumerable.DumpXLS();

            using (var fs = new FileStream("./excel.xls", FileMode.Create))
            {
                workbook.Write(fs);
            }
        }
    }
}
