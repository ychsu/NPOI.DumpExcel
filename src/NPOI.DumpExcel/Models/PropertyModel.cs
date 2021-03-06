﻿using NPOI.DumpExcel.Structs;
using System;
using System.Reflection;

namespace NPOI.DumpExcel.Models
{
    /// <summary>
    /// property model
    /// </summary>
    internal class PropertyModel
    {
        public int ColumnIndex { get; set; }

        /// <summary>
        /// column header name
        /// </summary>
        public string HeaderName { get; set; }

        /// <summary>
        /// property name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// column width
        /// </summary>
        public int Width { get; set; }

        /// <summary>
        /// cell format
        /// </summary>
        public CellFormatter CellFormatter { get; set; }

        /// <summary>
        /// property type
        /// </summary>
        public Type PropertyType { get; set; }

        /// <summary>
        /// property get method
        /// </summary>
        public MethodInfo GetMethod { get; set; }

        /// <summary>
        /// property set method
        /// </summary>
        public MethodInfo SetMethod { get; set; }
    }
}
