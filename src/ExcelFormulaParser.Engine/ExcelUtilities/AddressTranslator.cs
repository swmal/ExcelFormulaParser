using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    /// <summary>
    /// Handles translations from Spreadsheet addresses to 0-based numeric index.
    /// </summary>
    public class AddressTranslator
    {
        /// <summary>
        /// Translates an address in format "A1" to col- and rowindex.
        /// </summary>
        /// <param name="address"></param>
        /// <param name="col"></param>
        /// <param name="row"></param>
        public virtual void ToColAndRow(string address, out int col, out int row)
        {
            address = address.ToUpper();
            var alphaPart = GetAlphaPart(address);
            col = 0;
            var nLettersInAlphabet = 26;
            for (int x = 0; x < alphaPart.Length; x++)
            {
                var pos = alphaPart.Length - x - 1;
                var currentNumericValue = GetNumericAlphaValue(alphaPart[x]);
                col += (nLettersInAlphabet * pos * currentNumericValue);
                if (pos == 0)
                {
                    col += currentNumericValue;
                }
            }
            col--;
            row = GetIntPart(address) - 1;

        }

        private int GetNumericAlphaValue(char c)
        {
            return (int)c - 64;
        }

        private string GetAlphaPart(string address)
        {
            return Regex.Match(address, "[A-Z]+").Value;
        }

        private int GetIntPart(string address)
        {
            if (Regex.IsMatch(address, "[0-9]+"))
            {
                return int.Parse(Regex.Match(address, "[0-9]+").Value);
            }
            return -1;
        }
    }
}
