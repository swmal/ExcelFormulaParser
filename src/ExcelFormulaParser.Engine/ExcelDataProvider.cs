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
    public abstract class ExcelDataProvider
    {
        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="address">An Excel address</param>
        /// <returns>values from the required cells</returns>
        public abstract IEnumerable<object> GetRangeValues(string address);

        /// <summary>
        /// Returns values from the required range.
        /// </summary>
        /// <param name="address">An Excel address</param>
        /// <param name="excludeHiddenValues">If set to true, values from hidden cells should be excluded</param>
        /// <returns>values from the required cells</returns>
        public abstract IEnumerable<object> GetRangeValues(string address, bool excludeHiddenValues);
    }
}
