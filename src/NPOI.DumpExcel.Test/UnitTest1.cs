﻿using System;
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
                    DT = p % 2 == 0 ? new DateTimeOffset?(DateTimeOffset.Now.AddDays(p)) : null,
                    DT2 = DateTime.Now.AddDays(p),
                    Enum0 = Enum1.AAAAAAAAAAAAAAAAAAAAAA,
                    Enum1 = p % 2 == 0 ? CaseType.Report : new CaseType?(),
                    Name = $"Foo{p}",
                    SerId = p
                });

            var workbook = enumerable.DumpXLS();

            using (var fs = new FileStream("./DumpEnumerableToExcel.xls", FileMode.Create))
            {
                workbook.Write(fs);
            }
        }
        [TestMethod]
        public void DumpEnumerableToExcel2()
        {
            var enumerable = Enumerable.Range(1, 10000)
                .Select(p => new Member
                {
                    FirstName = "Foo",
                    LastName = "Foo",
                    Age = 18,
                    Birthday = DateTime.Now.AddYears(-18),
                    Gender = p % 2 == 0 ? Gender.Female : Gender.Male,
                    Height = 170,
                    IsMarried = p / 2 % 2 == 0,
                    UpdateOn = DateTime.Now
                });

            var workbook = enumerable.DumpXLSX();

            using (var fs = new FileStream("./DumpEnumerableToExcel2.xlsx", FileMode.Create))
            {
                workbook.Write(fs);
            }
        }

        [TestMethod]
        public void Dump2Sheets()
        {
            var parent = new Parent
            {
                Foos = Enumerable.Range(1, 10)
                .Select(p => new Foo
                {
                    DT = p % 2 == 0 ? new DateTimeOffset?(DateTimeOffset.Now.AddDays(p)) : null,
                    DT2 = DateTime.Now.AddDays(p),
                    Enum0 = Enum1.AAAAAAAAAAAAAAAAAAAAAA,
                    Enum1 = CaseType.Report,
                    Name = $"Foo{p}",
                    SerId = p
                }),
                Members = Enumerable.Range(1, 10000)
                .Select(p => new Member
                {
                    FirstName = "Foo",
                    LastName = "Foo",
                    Age = 18,
                    Birthday = DateTime.Now.AddYears(-18),
                    Gender = p % 2 == 0 ? Gender.Female : Gender.Male,
                    Height = 170,
                    IsMarried = p / 2 % 2 == 0,
                    UpdateOn = DateTime.Now
                })
            };

            var workbook = parent.Foos.DumpXLSX();
            workbook = parent.Members.DumpExcel(workbook);
            workbook = parent.Members.DumpExcel(workbook, "複製出來der");

            using (var fs = new FileStream("./DumpEnumerableToExcel3.xlsx", FileMode.Create))
            {
                workbook.Write(fs);
            }
        }
    }
}
