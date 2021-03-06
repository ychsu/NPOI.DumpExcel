﻿using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI.DumpExcel
{
    public static class EnumerableExtension
    {
        private static IDictionary<string, Type> _dumpServices;

        static EnumerableExtension()
        {
            _dumpServices = new Dictionary<string, Type>();
        }

        public static IWorkbook DumpExcel<T>(this IEnumerable<T> enumerable, 
            IWorkbook workbook, string sheetName = null)
        {
            var type = typeof(T);
            var dumpServType = FindService(type);

            var dumpServ = Activator.CreateInstance(dumpServType, new object[] { workbook, sheetName }) as IDumpService<T>;

            return dumpServ.DumpWorkbook(enumerable);
        }

        /// <summary>
        /// Dump enumerable 成workbook
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable"></param>
        /// <param name="excelType"></param>
        /// <returns></returns>
        public static IWorkbook DumpExcel<T>(this IEnumerable<T> enumerable, ExcelType excelType, string sheetName = null)
        {
            var workbook = excelType == ExcelType.XLSX ?
                new NPOI.XSSF.UserModel.XSSFWorkbook() as IWorkbook :
                new NPOI.HSSF.UserModel.HSSFWorkbook();

            return enumerable.DumpExcel(workbook, sheetName);
        }


        /// <summary>
        /// Dump enumerable 成workbook (xls)
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="enumerable"></param>
        /// <param name="excelType"></param>
        /// <returns></returns>
        public static IWorkbook DumpXLS<T>(this IEnumerable<T> enumerable, string sheetName = null) { return DumpExcel(enumerable, ExcelType.XLS, sheetName); }

        public static IWorkbook DumpXLSX<T>(this IEnumerable<T> enumerable, string sheetName = null) { return DumpExcel(enumerable, ExcelType.XLSX, sheetName); }

        public static Type FindService(Type type)
        {
            var key = $"{type.FullName}, {type.Assembly.FullName}";
            if (_dumpServices.ContainsKey(key) == true)
            {
                return _dumpServices[key];
            }

            var service = DumpUtil.CreateService(type);
            _dumpServices.Add(key, service);
            return service;
        }
    }
}
