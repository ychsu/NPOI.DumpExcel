using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace NPOI.DumpExcel
{
    internal static class TypeExtensions
    {
        public static T GetCustomAttribute<T>(this MemberInfo type)
        {
            return (T)type.GetCustomAttributes(typeof(T), true).FirstOrDefault();
        }
    }
}
