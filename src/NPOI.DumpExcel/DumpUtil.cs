using NPOI.DumpExcel.Attributes;
using NPOI.DumpExcel.Models;
using NPOI.DumpExcel.Structs;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Threading;

namespace NPOI.DumpExcel
{
    /// <summary>
    /// dump utility
    /// </summary>
    public class DumpUtil
    {
        private static AssemblyBuilder assemblyBuilder;
        private static ModuleBuilder moduleBuilder;
        private static MethodInfo createRow;
        private static MethodInfo physicalNumberOfRows;
        private static MethodInfo createCell;
        private static MethodInfo[] setCellValues; // 會放在static是犧牲記憶體求取效能
        private static MethodInfo getLocalDateTime;
        static DumpUtil()
        {
            var sheetType = typeof(ISheet);
            var rowType = typeof(IRow);
            var cellType = typeof(ICell);
            // 定義assembly
            assemblyBuilder = Thread.GetDomain().DefineDynamicAssembly(new AssemblyName { Name = "YC.NPOI.DumpExcel" }, AssemblyBuilderAccess.RunAndSave);
            // 定義module
            moduleBuilder = assemblyBuilder.DefineDynamicModule("YC.NPOI.DumpExcel", "YC.NPOI.DumpExcel.dll");
            createRow = sheetType.GetMethod("CreateRow");
            physicalNumberOfRows = sheetType.GetMethod("get_PhysicalNumberOfRows");
            createCell = rowType.GetMethod("CreateCell", new Type[] { typeof(int) });
            setCellValues = new MethodInfo[]
            {
                cellType.GetMethod("SetCellValue", new Type[] { typeof(double) }),
                cellType.GetMethod("SetCellValue", new Type[] { typeof(DateTime) }),
                cellType.GetMethod("SetCellValue", new Type[] { typeof(string) }),
                cellType.GetMethod("SetCellValue", new Type[] { typeof(bool) })
            };
            getLocalDateTime = typeof(DateTimeOffset).GetMethod("get_LocalDateTime");
        }

        /// <summary>
        /// 建立DumpService
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        internal static Type CreateService(Type type)
        {
            var sheetname = type.GetCustomAttribute<SheetAttribute>()?.Name ?? type.Name;
            var baseType = typeof(DumpServiceBase<>).MakeGenericType(type);
            var @interface = typeof(IDumpService<>).MakeGenericType(type);
            var className = $"{type.Name}DumpService";
            if (string.IsNullOrEmpty(type.Namespace) == false)
            {
                className = $"{type.Namespace}.{className}";
            }
            var typeBuilder = moduleBuilder.DefineType(className,
                TypeAttributes.Sealed | TypeAttributes.Public | TypeAttributes.Class,
                baseType,
                new Type[] { @interface });


            var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Select(t => new
                {
                    prop = t,
                    excelAttr = t.GetCustomAttribute<ExcelColumnAttribute>(),
                    displayAttr = t.GetCustomAttribute<DisplayAttribute>()
                })
                .Where(t => t.excelAttr != null)
                .OrderBy(t => t.excelAttr.Order)
                .Select((t, idx) => new PropertyModel
                {
                    ColumnIndex = idx,
                    GetMethod = t.prop.GetGetMethod(),
                    SetMethod = t.prop.GetSetMethod(),
                    HeaderName = t.displayAttr?.GetName() ?? t.prop.Name,
                    Width = t.excelAttr.Width,
                    CellFormatter = t.excelAttr.CellFormatter,
                    Name = t.prop.Name,
                    PropertyType = t.prop.PropertyType
                })
                .ToList();

            #region filed 
            var sheet = typeBuilder.DefineField("sheet", typeof(ISheet), FieldAttributes.Private);
            #endregion

            var emitSetColumnsStyle = EmitSetColumnsStyle(typeBuilder, sheet, properties);

            var emitSetColumnsWidth = EmitSetColumnsWidth(typeBuilder, sheet, properties);

            EmitConstructor(sheetname, baseType, typeBuilder, new MethodBuilder[]
            {
                emitSetColumnsStyle,
                emitSetColumnsWidth
            }, properties, sheet);

            EmitCreateHeaderRow(typeBuilder, baseType, sheet, properties);

            EmitCreateRow(typeBuilder, baseType, sheet, properties);

            var serviceType = typeBuilder.CreateType();

#if SAVEASSEMBLY
            assemblyBuilder.Save("YC.NPOI.DumpExcel.dll");
#endif
            return serviceType;
        }

        /// <summary>
        /// 建立type的constructor
        /// </summary>
        /// <param name="sheetname"></param>
        /// <param name="baseType"></param>
        /// <param name="typeBuilder"></param>
        /// <param name="properties"></param>
        /// <param name="sheet"></param>
        /// <param name="styles"></param>
        private static void EmitConstructor(string sheetname, Type baseType, TypeBuilder typeBuilder, IEnumerable<MethodBuilder> methods, IList<PropertyModel> properties, FieldBuilder sheet)
        {
            var constructParameterType = typeof(IWorkbook);
            var constructBuilder = typeBuilder.DefineConstructor(
                MethodAttributes.Public | MethodAttributes.PrivateScope,
                CallingConventions.Standard,
                new Type[] { constructParameterType });
            if (constructBuilder != null)
            {
                var il = constructBuilder.GetILGenerator();
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldarg_1);
                il.Emit(OpCodes.Call, baseType.GetConstructor(new Type[] { constructParameterType }));
                il.Emit(OpCodes.Nop);
                il.Emit(OpCodes.Nop);
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldarg_1);
                il.Emit(OpCodes.Ldstr, sheetname);
                il.Emit(OpCodes.Callvirt, typeof(IWorkbook).GetMethod("CreateSheet", new Type[] { typeof(string) }));
                il.Emit(OpCodes.Stfld, sheet);

                foreach (var method in methods)
                {
                    il.Emit(OpCodes.Ldarg_0);
                    il.Emit(OpCodes.Call, method);
                }

                il.Emit(OpCodes.Nop);
                il.Emit(OpCodes.Ret);
            }
        }

        private static MethodBuilder EmitSetColumnsStyle(TypeBuilder builder, FieldBuilder sheet, IList<PropertyModel> properties)
        {
            var methodBuilder = builder.DefineMethod("SetColumnsStyle",
                MethodAttributes.Private | MethodAttributes.PrivateScope);

            var il = methodBuilder.GetILGenerator();

            #region local variables
            var dataFormat = il.DeclareLocal(typeof(IDataFormat));
            var style = il.DeclareLocal(typeof(ICellStyle));
            #endregion

            var method_CreateDataFormat = typeof(IWorkbook).GetMethod("CreateDataFormat", new Type[] { });
            var method_GetFormat = typeof(IDataFormat).GetMethod("GetFormat", new Type[] { typeof(string) });
            var method_CreateCellStyle = typeof(IWorkbook).GetMethod("CreateCellStyle", Type.EmptyTypes);
            var method_SetDefaultColumnStyle = typeof(ISheet).GetMethod("SetDefaultColumnStyle");
            var prop_Workbook = sheet.FieldType.GetProperty("Workbook");
            var prop_DataFormat = style.LocalType.GetProperty("DataFormat");
            var prop_Alignment = style.LocalType.GetProperty("Alignment");
            var prop_VerticalAlignment = style.LocalType.GetProperty("VerticalAlignment");

            il.Emit(OpCodes.Nop);

            #region CreateDataFormat;
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Callvirt, prop_Workbook.GetGetMethod());
            il.Emit(OpCodes.Callvirt, method_CreateDataFormat);
            il.Emit(OpCodes.Stloc, dataFormat);
            #endregion

            var groupByFormat = properties.GroupBy(p => p.CellFormatter)
                .Select(g => new
                {
                    CellFormat = g.Key,
                    Items = g
                });
            var formats = properties.Select(p => p.CellFormatter).Distinct().ToArray();
            foreach (var format in groupByFormat)
            {
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldfld, sheet);
                il.Emit(OpCodes.Callvirt, prop_Workbook.GetGetMethod());
                il.Emit(OpCodes.Callvirt, method_CreateCellStyle);
                il.Emit(OpCodes.Stloc, style);

                if (string.IsNullOrEmpty(format.CellFormat.Format) == false)
                {
                    il.Emit(OpCodes.Ldloc, style);
                    il.Emit(OpCodes.Ldloc, dataFormat);
                    il.Emit(OpCodes.Ldstr, format.CellFormat.Format);
                    il.Emit(OpCodes.Callvirt, method_GetFormat);
                    il.Emit(OpCodes.Callvirt, prop_DataFormat.GetSetMethod());
                }

                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Ldc_I4, (int)format.CellFormat.Alignment);
                il.Emit(OpCodes.Callvirt, prop_Alignment.GetSetMethod());

                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Ldc_I4, (int)format.CellFormat.VerticalAlign);
                il.Emit(OpCodes.Callvirt, prop_VerticalAlignment.GetSetMethod());

                // set columns
                foreach (var prop in format.Items)
                {
                    il.Emit(OpCodes.Ldarg_0);
                    il.Emit(OpCodes.Ldfld, sheet);
                    il.Emit(OpCodes.Ldc_I4, prop.ColumnIndex);
                    il.Emit(OpCodes.Ldloc, style);
                    il.Emit(OpCodes.Callvirt, method_SetDefaultColumnStyle);
                }
            }

            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Ret);

            return methodBuilder;
        }

        /// <summary>
        /// 設定所有欄位欄寬 (如果有設定)
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="sheet"></param>
        /// <param name="properties"></param>
        private static MethodBuilder EmitSetColumnsWidth(TypeBuilder builder, FieldBuilder sheet, IList<PropertyModel> properties)
        {
            var methodBuilder = builder.DefineMethod("SetColumnsWidth",
                MethodAttributes.Private | MethodAttributes.PrivateScope);

            var il = methodBuilder.GetILGenerator();

            il.Emit(OpCodes.Nop);

            for (int i = 0; i < properties.Count(); i++)
            {
                var prop = properties.ElementAt(i);
                if (prop.Width > 0)
                {
                    il.Emit(OpCodes.Ldarg_0);
                    il.Emit(OpCodes.Ldfld, sheet);
                    il.Emit(OpCodes.Ldc_I4, i);
                    il.Emit(OpCodes.Ldc_I4, (int)((prop.Width + 0.72) * 256));
                    il.Emit(OpCodes.Callvirt, sheet.FieldType.GetMethod("SetColumnWidth"));
                }
            }

            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Ret);

            return methodBuilder;
        }


        /// <summary>
        /// 建立 method CreateHeaderRow (因繼承DumpServiceBase, 所以實作)
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="baseType"></param>
        /// <param name="sheet"></param>
        /// <param name="properties"></param>
        private static void EmitCreateHeaderRow(TypeBuilder builder, Type baseType, FieldBuilder sheet, IEnumerable<PropertyModel> properties)
        {
            var methodBuilder = builder.DefineMethod("CreateHeaderRow",
                MethodAttributes.PrivateScope | MethodAttributes.Family | MethodAttributes.ReuseSlot | MethodAttributes.Virtual | MethodAttributes.HideBySig);

            var il = methodBuilder.GetILGenerator();
            var row = il.DeclareLocal(typeof(IRow));  // 定義一個local變數
            il.Emit(OpCodes.Nop);

            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Ldc_I4_0);
            il.Emit(OpCodes.Callvirt, createRow);
            il.Emit(OpCodes.Stloc, row); // loc_0 當做row

            for (var i = 0; i < properties.Count(); i++)
            {
                var prop = properties.ElementAt(i);
                var cell = il.DeclareLocal(typeof(ICell));

                il.Emit(OpCodes.Ldloc, row);
                il.Emit(OpCodes.Ldc_I4, i);
                il.Emit(OpCodes.Callvirt, createCell);
                il.Emit(OpCodes.Stloc, cell);
                il.Emit(OpCodes.Ldloc, cell);
                il.Emit(OpCodes.Ldstr, prop.HeaderName);
                il.Emit(OpCodes.Callvirt, setCellValues[2]);
            }

            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Ret);
            builder.DefineMethodOverride(methodBuilder,
                baseType.GetMethod("CreateHeaderRow", BindingFlags.NonPublic | BindingFlags.Instance));
        }

        /// <summary>
        /// 建立method CreateRow (因繼承DumpServiceBase, 所以實作)
        /// </summary>
        /// <param name="builder"></param>
        /// <param name="baseType"></param>
        /// <param name="sheet"></param>
        /// <param name="properties"></param>
        private static void EmitCreateRow(TypeBuilder builder, Type baseType, FieldBuilder sheet, IEnumerable<PropertyModel> properties)
        {
            var type = baseType.GetGenericArguments()[0];
            var methodBuilder = builder.DefineMethod("CreateRow",
                MethodAttributes.PrivateScope | MethodAttributes.Family | MethodAttributes.ReuseSlot | MethodAttributes.Virtual | MethodAttributes.HideBySig,
                null,
                new Type[] { type });

            var il = methodBuilder.GetILGenerator();
            var row = il.DeclareLocal(typeof(IRow));  // 定義一個local變數
            il.Emit(OpCodes.Nop);

            var rowNum = il.DeclareLocal(typeof(int));
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Callvirt, physicalNumberOfRows);
            il.Emit(OpCodes.Stloc, rowNum);
            il.Emit(OpCodes.Nop);

            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Ldloc, rowNum);
            il.Emit(OpCodes.Callvirt, createRow);
            il.Emit(OpCodes.Stloc, row); // loc_0 當做row
            il.Emit(OpCodes.Nop);

            var cell = il.DeclareLocal(typeof(ICell));
            foreach (var property in properties)
            {
                CreateCell(il, row, cell, property);
            }

            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Ret);
            builder.DefineMethodOverride(methodBuilder,
                baseType.GetMethod("CreateRow", BindingFlags.NonPublic | BindingFlags.Instance));
        }

        private static void CreateCell(ILGenerator il, LocalBuilder row, LocalBuilder cell, PropertyModel property)
        {
            il.Emit(OpCodes.Ldloc, row);
            il.Emit(OpCodes.Ldc_I4, property.ColumnIndex);
            il.Emit(OpCodes.Callvirt, createCell);
            il.Emit(OpCodes.Stloc, cell);

            #region NPOI 2.3前 匯出xlsx會有吃不到 DefaultColumnStyle 的問題
            il.Emit(OpCodes.Ldloc, cell);
            il.Emit(OpCodes.Ldloc, row);
            il.Emit(OpCodes.Callvirt, typeof(IRow).GetProperty("Sheet").GetGetMethod());
            il.Emit(OpCodes.Ldc_I4, property.ColumnIndex);
            il.Emit(OpCodes.Callvirt, typeof(ISheet).GetMethod("GetColumnStyle"));
            il.Emit(OpCodes.Callvirt, typeof(ICell).GetProperty("CellStyle").GetSetMethod());
            #endregion
            il.Emit(OpCodes.Nop);

            var val = il.DeclareLocal(property.PropertyType);
            il.Emit(OpCodes.Ldarg_1);
            il.Emit(OpCodes.Callvirt, property.GetMethod);
            il.Emit(OpCodes.Stloc, val);

            var underlayingType = Nullable.GetUnderlyingType(property.PropertyType);
            if (underlayingType != null)
            {
                var realValue = il.DeclareLocal(underlayingType);
                var inequal = il.DeclareLocal(typeof(bool));
                var lbl = il.DefineLabel();
                il.Emit(OpCodes.Ldloca, val);
                il.Emit(OpCodes.Callvirt, property.PropertyType.GetProperty("HasValue").GetGetMethod());
                il.Emit(OpCodes.Brfalse_S, lbl);
                il.Emit(OpCodes.Nop);

                il.Emit(OpCodes.Ldloca, val);
                il.Emit(OpCodes.Call, property.PropertyType.GetMethod("GetValueOrDefault", Type.EmptyTypes));
                il.Emit(OpCodes.Stloc, realValue);

                SetCellValue(il, cell, realValue);

                il.MarkLabel(lbl);
                il.Emit(OpCodes.Nop);
            }
            else
            {
                SetCellValue(il, cell, val);
            }
            il.Emit(OpCodes.Nop);
        }

        private static void SetCellValue(ILGenerator il, LocalBuilder cell, LocalBuilder val, string typeName = null)
        {
            il.Emit(OpCodes.Ldloc, cell);
                    il.Emit(OpCodes.Ldloc, val);
            switch (typeName ?? val.LocalType.Name)
            {
                case "Decimal":
                    il.Emit(OpCodes.Call,
                        typeof(Decimal).GetMethod("ToDouble", BindingFlags.Static | BindingFlags.Public));
                    SetNumericCellValue(il, cell, val);
                    break;
                case "Int16":
                case "Int32":
                case "Int64":
                case "UInt16":
                case "UInt32":
                case "UInt64":
                case "Short":
                case "Double":
                case "Single":
                    SetNumericCellValue(il, cell, val);
                    break;
                case nameof(DateTimeOffset):
                    il.Emit(OpCodes.Pop);
                    il.Emit(OpCodes.Ldloca_S, val);
                    il.Emit(OpCodes.Call, typeof(DateTimeOffset).GetProperty("LocalDateTime").GetGetMethod());
                    SetDateTimeCellValue(il, cell, val);
                    break;
                case "DateTime":
                    SetDateTimeCellValue(il, cell, val);
                    break;
                case "Boolean":
                    SetBooleanCellValue(il, cell, val);
                    break;
                default:
                    SetStringCellValue(il, cell, val);
                    break;
            }
        }

        private static void SetNumericCellValue(ILGenerator il, LocalBuilder cell, LocalBuilder val)
        {
            il.Emit(OpCodes.Conv_R8); // convert to double;
            il.Emit(OpCodes.Callvirt, setCellValues[0]);
        }

        private static void SetDateTimeCellValue(ILGenerator il, LocalBuilder cell, LocalBuilder val)
        {
            il.Emit(OpCodes.Callvirt, setCellValues[1]);
        }

        private static void SetStringCellValue(ILGenerator il, LocalBuilder cell, LocalBuilder val)
        {
            if (val.LocalType != typeof(string))
            {
                il.Emit(OpCodes.Box, val.LocalType);
                il.Emit(OpCodes.Call,
                    typeof(System.Convert).GetMethod("ToString",
                        BindingFlags.Public | BindingFlags.Static,
                        null,
                        CallingConventions.Any,
                        new Type[] { typeof(object) },
                        null));
            }
            il.Emit(OpCodes.Callvirt, setCellValues[2]);
        }

        private static void SetBooleanCellValue(ILGenerator il, LocalBuilder cell, LocalBuilder val)
        {
            il.Emit(OpCodes.Callvirt, setCellValues[3]);
        }
    }
}
