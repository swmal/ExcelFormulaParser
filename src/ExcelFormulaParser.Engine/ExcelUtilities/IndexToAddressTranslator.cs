using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class IndexToAddressTranslator
    {
        const int MaxAlphabetIndex = 25;
        const int NLettersInAlphabet = 26;

        public string ToAddress(int col, int row)
        {
            if (col <= MaxAlphabetIndex)
            {
                return string.Concat(IntToChar(col), row + 1);
            }
            else if (col < (Math.Pow(NLettersInAlphabet, 2) + NLettersInAlphabet))
            {
                var firstChar = col / NLettersInAlphabet - 1;
                var secondChar = col % NLettersInAlphabet;
                return string.Concat(IntToChar(firstChar), IntToChar(secondChar), row + 1);
            }
            else if(col < (Math.Pow(NLettersInAlphabet, 3) + NLettersInAlphabet))
            {
                var x = NLettersInAlphabet * NLettersInAlphabet;
                var rest = col - x;
                var firstChar = col / x - 1;
                var secondChar = rest / NLettersInAlphabet - 1;
                var thirdChar = rest % NLettersInAlphabet;
                return string.Concat(IntToChar(firstChar), IntToChar(secondChar), IntToChar(thirdChar), row + 1);
            }
            throw new InvalidOperationException("ExcelFormulaParser does not the supplied number of columns " + col);
        }

        private char IntToChar(int i)
        {
            return (char)(i + 65);
        }
    }
}
