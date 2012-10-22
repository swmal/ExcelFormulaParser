using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine
{
    /// <summary>
    /// This class should be implemented to be able to deliver excel data
    /// to the formula parser.
    /// </summary>
    public abstract class ExcelDataProvider : IDisposable
    {
        /// <summary>
        /// Returns all formulas on a worksheet
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public abstract IDictionary<string, string> GetWorksheetFormulas(string sheetName);

        /// <summary>
        /// Returns all formulas in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract IDictionary<string, string> GetWorkbookFormulas();

        /// <summary>
        /// Returns all defined names in a workbook
        /// </summary>
        /// <returns></returns>
        public abstract IDictionary<string, object> GetWorkbookNameValues();
        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="address">An Excel address</param>
        /// <returns>values from the required cells</returns>
        public abstract IEnumerable<ExcelCell> GetRangeValues(string address);

        /// <summary>
        /// Returs a matrix for lookup functions like VLOOKUP and HLOOKUP.
        /// </summary>
        /// <param name="address"></param>
        /// <returns></returns>
        public abstract IDictionary<int, IList<ExcelCell>> GetLookupArray(string address);

        /// <summary>
        /// Sets the value on the cell
        /// </summary>
        /// <param name="address"></param>
        /// <param name="value"></param>
        public abstract void SetCellValue(string address, object value);

        /// <summary>
        /// Use this method to free unmanaged resources.
        /// </summary>
        public abstract void Dispose();

        /// <summary>
        /// Max number of columns in a worksheet that the Excel data provider can handle.
        /// </summary>
        public abstract int ExcelMaxColumns { get; }

        /// <summary>
        /// Max number of rows in a worksheet that the Excel data provider can handle
        /// </summary>
        public abstract int ExcelMaxRows { get; }
    }
}
