using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.Text
{
    public class Replace : VBAFunction
    {
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 4);
            var oldText = ArgToString(arguments, 0);
            var startPos = ArgToInt(arguments, 1);
            var nCharsToReplace = ArgToInt(arguments, 2);
            var newText = ArgToString(arguments, 3);
            var firstPart = GetFirstPart(oldText, startPos);
            var lastPart = GetLastPart(oldText, startPos, nCharsToReplace);
            var result = string.Concat(firstPart, newText, lastPart);
            return new CompileResult(result, DataType.String);
        }

        private string GetFirstPart(string text, int startPos)
        {
            return text.Substring(0, startPos - 1);
        }

        private string GetLastPart(string text, int startPos, int nCharactersToReplace)
        {
            int startIx = startPos -1;
            startIx += nCharactersToReplace;
            return text.Substring(startIx, text.Length - startIx);
        }
    }
}
