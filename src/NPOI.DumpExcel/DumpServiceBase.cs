using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.IO;

namespace NPOI.DumpExcel
{
    /// <summary>
    /// dumpservicebase
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public abstract class DumpServiceBase<T> : IDumpService<T>
    {
        protected IWorkbook workbook;

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="workbook">workbook</param>
        public DumpServiceBase(IWorkbook workbook)
        {
            this.workbook = workbook;
        }

        /// <summary>
        /// create header row
        /// </summary>
        protected abstract void CreateHeaderRow();

        /// <summary>
        /// create data row
        /// </summary>
        /// <param name="entity">data</param>
        protected abstract void CreateRow(T entity);

        /// <summary>
        /// dump enumerable to excel workbook
        /// </summary>
        /// <param name="enumerable"></param>
        /// <returns></returns>
        public virtual IWorkbook DumpWorkbook(IEnumerable<T> enumerable)
        {
            CreateHeaderRow();
            foreach (var item in enumerable)
            {
                CreateRow(item);
            }

            return this.workbook;
        }
    }
}
