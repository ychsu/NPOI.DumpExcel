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

#if DEBUG
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
                var constructIl = constructBuilder.GetILGenerator();
                constructIl.Emit(OpCodes.Ldarg_0);
                constructIl.Emit(OpCodes.Ldarg_1);
                constructIl.Emit(OpCodes.Call, baseType.GetConstructor(new Type[] { constructParameterType }));
                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Ldarg_0);
                constructIl.Emit(OpCodes.Ldarg_1);
                constructIl.Emit(OpCodes.Ldstr, sheetname);
                constructIl.Emit(OpCodes.Call, typeof(IWorkbook).GetMethod("CreateSheet", new Type[] { typeof(string) }));
                constructIl.Emit(OpCodes.Stfld, sheet);

                foreach (var method in methods)
                {
                    constructIl.Emit(OpCodes.Ldarg_0);
                    constructIl.Emit(OpCodes.Call, method);
                }

                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Ret);
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
            il.Emit(OpCodes.Call, method_CreateDataFormat);
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
                il.Emit(OpCodes.Call, method_CreateCellStyle);
                il.Emit(OpCodes.Stloc, style);

                if (string.IsNullOrEmpty(format.CellFormat.Format) == false)
                {
                    il.Emit(OpCodes.Ldloc, style);
                    il.Emit(OpCodes.Ldloc, dataFormat);
                    il.Emit(OpCodes.Ldstr, format.CellFormat.Format);
                    il.Emit(OpCodes.Call, method_GetFormat);
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
                    il.Emit(OpCodes.Call, method_SetDefaultColumnStyle);
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
                il.Emit(OpCodes.Call, setCellValues[2]);
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
            var methodBuilder = builder.DefineMethod("CreateRow",
                MethodAttributes.PrivateScope | MethodAttributes.Family | MethodAttributes.ReuseSlot | MethodAttributes.Virtual | MethodAttributes.HideBySig,
                null,
                new Type[] { baseType.GetGenericArguments()[0] });

            var il = methodBuilder.GetILGenerator();
            var row = il.DeclareLocal(typeof(IRow));  // 定義一個local變數
            il.Emit(OpCodes.Nop);

            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Ldarg_0);
            il.Emit(OpCodes.Ldfld, sheet);
            il.Emit(OpCodes.Call, physicalNumberOfRows);
            il.Emit(OpCodes.Callvirt, createRow);
            il.Emit(OpCodes.Stloc, row); // loc_0 當做row

            for (var i = 0; i < properties.Count(); i++)
            {
                var property = properties.ElementAt(i);
                var cell = il.DeclareLocal(typeof(ICell));
                il.Emit(OpCodes.Ldloc, row);
                il.Emit(OpCodes.Ldc_I4, i);
                il.Emit(OpCodes.Callvirt, createCell);
                il.Emit(OpCodes.Stloc, cell);
                SetCellValue(property, il, cell);
            }

            il.Emit(OpCodes.Nop);
            il.Emit(OpCodes.Ret);
            builder.DefineMethodOverride(methodBuilder,
                baseType.GetMethod("CreateRow", BindingFlags.NonPublic | BindingFlags.Instance));
        }

        /// <summary>
        /// 設定CellValue
        /// </summary>
        /// <param name="property"></param>
        /// <param name="il"></param>
        /// <param name="cell"></param>
        private static void SetCellValue(PropertyModel property, ILGenerator il, LocalBuilder cell)
        {
            il.Emit(OpCodes.Ldloc, cell);
            il.Emit(OpCodes.Ldarg_1);
            il.Emit(OpCodes.Callvirt, property.GetMethod);
            switch (property.PropertyType.Name)
            {
                case "Int16":
                case "Int32":
                case "Int64":
                case "UInt16":
                case "UInt32":
                case "UInt64":
                case "Short":
                case "Double":
                case "Single":
                    SetNumericCellValue(property, il, cell);
                    break;
                case "DateTime":
                    SetDateTimeCellValue(property, il, cell);
                    break;
                case "Boolean":
                    SetBooleanCellValue(property, il, cell);
                    break;
                default:
                    SetStringCellValue(property, il, cell);
                    break;
            }
        }

        //private static Action<PropertyInfo, ILGenerator, ICell> GetSetCellValueMethod(Type propertyType)
        //{
        //    if (propertyType.IsNested == false)
        //    {
        //        if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
        //        {
        //            return GetSetCellValueMethod(propertyType.GenericTypeArguments[0]);
        //        }

        //        if (propertyType.IsPrimitive == false)
        //        {

        //        }
        //    }
        //    return null;
        //}

        private static void SetNumericCellValue(PropertyModel property, ILGenerator il, LocalBuilder cell)
        {
            il.Emit(OpCodes.Conv_R8); // convert to double;
            il.Emit(OpCodes.Call, setCellValues[0]);
        }

        private static void SetDateTimeCellValue(PropertyModel property, ILGenerator il, LocalBuilder cell)
        {
            il.Emit(OpCodes.Call, setCellValues[1]);
        }

        private static void SetStringCellValue(PropertyModel property, ILGenerator il, LocalBuilder cell)
        {
            if (property.PropertyType != typeof(string))
            {
                il.Emit(OpCodes.Box, property.PropertyType);
                il.Emit(OpCodes.Call,
                    typeof(System.Convert).GetMethod("ToString",
                        BindingFlags.Public | BindingFlags.Static,
                        null,
                        CallingConventions.Any,
                        new Type[] { typeof(object) },
                        null));
            }
            il.Emit(OpCodes.Call, setCellValues[2]);
        }

        private static void SetBooleanCellValue(PropertyModel property, ILGenerator il, LocalBuilder cell)
        {
            il.Emit(OpCodes.Call, setCellValues[3]);
        }
    }
}
