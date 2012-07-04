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
        /// Returns values from the required range.
        /// </summary>
        /// <param name="address">An Excel address</param>
        /// <returns>values from the required cells</returns>
        public abstract IEnumerable<ExcelDataItem> GetRangeValues(string address);

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
