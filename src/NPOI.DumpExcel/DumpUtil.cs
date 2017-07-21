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
                .Select(t => new PropertyModel
                {
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
            var stylesType = typeof(Dictionary<int, ICellStyle>);
            var styles = typeBuilder.DefineField("styles", stylesType, FieldAttributes.Private);
            #endregion

            EmitConstructor(sheetname, baseType, typeBuilder, properties, sheet, styles);

            EmitCreateHeaderRow(typeBuilder, baseType, sheet, properties);

            EmitCreateRow(typeBuilder, baseType, sheet, styles, properties);

            var serviceType = typeBuilder.CreateType();


            assemblyBuilder.Save("YC.NPOI.DumpExcel.dll");
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
        private static void EmitConstructor(string sheetname, Type baseType, TypeBuilder typeBuilder, IList<PropertyModel> properties, FieldBuilder sheet, FieldBuilder styles)
        {
            var constructParameterType = typeof(IWorkbook);
            var constructBuilder = typeBuilder.DefineConstructor(
                MethodAttributes.Public,
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

                var format = constructIl.DeclareLocal(typeof(IDataFormat));
                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Ldarg_1);
                var createDataFormat = typeof(IWorkbook).GetMethod("CreateDataFormat", new Type[] { });
                constructIl.Emit(OpCodes.Call, createDataFormat);
                constructIl.Emit(OpCodes.Stloc, format);

                constructIl.Emit(OpCodes.Ldarg_0);
                constructIl.Emit(OpCodes.Newobj, styles.FieldType.GetConstructor(new Type[] { }));
                constructIl.Emit(OpCodes.Stfld, styles);

                SetFormatters(constructIl, styles, format, properties);

                SetWidth(constructIl, sheet, properties);

                constructIl.Emit(OpCodes.Nop);
                constructIl.Emit(OpCodes.Ret);
            }
        }

        /// <summary>
        /// 設定所有欄位欄寬 (如果有設定)
        /// </summary>
        /// <param name="il"></param>
        /// <param name="sheet"></param>
        /// <param name="properties"></param>
        private static void SetWidth(ILGenerator il, FieldBuilder sheet, IList<PropertyModel> properties)
        {
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
        }

        /// <summary>
        /// 建立format並且寫到field中 
        /// </summary>
        /// <param name="il"></param>
        /// <param name="styles"></param>
        /// <param name="format"></param>
        /// <param name="properties"></param>
        private static void SetFormatters(ILGenerator il,
            FieldBuilder styles,
            LocalBuilder format,
            IList<PropertyModel> properties)
        {
            var getFormat = typeof(IDataFormat).GetMethod("GetFormat", new Type[] { typeof(string) });
            var createCellStyle = typeof(IWorkbook).GetMethod("CreateCellStyle", Type.EmptyTypes);
            var cellStyleType = typeof(ICellStyle);
            var dataFormat = cellStyleType.GetProperty("DataFormat");
            var alignment = cellStyleType.GetProperty("Alignment");
            var verticalAlignment = cellStyleType.GetProperty("VerticalAlignment");

            var formats = properties.Select(p => p.CellFormatter).Distinct().ToArray();
            for (var i = 0; i < formats.Length; i++)
            {
                foreach (var prop in properties.Where(p => p.CellFormatter.Equals(formats[i]) == true))
                {
                    prop.FormatId = i;
                }
                var f = formats[i];

                var style = il.DeclareLocal(typeof(ICellStyle));
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldarg_1);
                il.Emit(OpCodes.Call, createCellStyle);
                il.Emit(OpCodes.Stloc, style);

                if (string.IsNullOrEmpty(f.Format) == false)
                {
                    il.Emit(OpCodes.Ldloc, style);
                    il.Emit(OpCodes.Ldloc, format);
                    il.Emit(OpCodes.Ldstr, f.Format);
                    il.Emit(OpCodes.Call, getFormat);
                    il.Emit(OpCodes.Callvirt, dataFormat.GetSetMethod());
                }

                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Ldc_I4, (int)f.Alignment);
                il.Emit(OpCodes.Callvirt, alignment.GetSetMethod());

                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Ldc_I4, (int)f.VerticalAlign);
                il.Emit(OpCodes.Callvirt, verticalAlignment.GetSetMethod());

                il.Emit(OpCodes.Ldfld, styles);
                il.Emit(OpCodes.Ldc_I4, i);
                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Callvirt, styles.FieldType.GetMethod("Add"));
            }
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
        /// <param name="styles"></param>
        /// <param name="properties"></param>
        private static void EmitCreateRow(TypeBuilder builder, Type baseType, FieldBuilder sheet, FieldBuilder styles, IEnumerable<PropertyModel> properties)
        {
            var styleDicType = styles.FieldType;
            var get_Item = styleDicType.GetMethod("get_Item");
            var set_CellStyle = typeof(ICell).GetMethod("set_CellStyle");

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
                var style = il.DeclareLocal(typeof(ICellStyle));
                il.Emit(OpCodes.Ldarg_0);
                il.Emit(OpCodes.Ldfld, styles);
                il.Emit(OpCodes.Ldc_I4, property.FormatId);
                il.Emit(OpCodes.Callvirt, get_Item);
                il.Emit(OpCodes.Stloc, style);
                il.Emit(OpCodes.Ldloc, cell);
                il.Emit(OpCodes.Ldloc, style);
                il.Emit(OpCodes.Callvirt, set_CellStyle);
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
