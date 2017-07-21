using System;

namespace NPOI.DumpExcel.Attributes
{
    /// <summary>
    /// worksheet
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class SheetAttribute : Attribute
    {
        /// <summary>
        /// worksheet 
        /// </summary>
        /// <param name="Name">sheet name</param>
        public SheetAttribute(string Name = null)
        {
            this.Name = Name;
        }

        /// <summary>
        /// sheet name
        /// </summary>
        public string Name { get; set; }
    }
}