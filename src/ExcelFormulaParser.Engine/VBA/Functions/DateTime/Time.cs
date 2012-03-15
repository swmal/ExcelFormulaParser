using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.VBA.Functions.DateTime
{
    public class Time : TimeBaseFunction
    {
        public Time()
            : this(new TimeStringParser())
        {

        }

        public Time(TimeStringParser timeStringParser)
        {
            _timeStringParser = timeStringParser;
        }

        private readonly TimeStringParser _timeStringParser;
        
        public override CompileResult Execute(IEnumerable<object> arguments)
        {
            ValidateArguments(arguments, 1);
            var firstArg = arguments.ElementAt(0).ToString();
            if(arguments.Count() == 1 && _timeStringParser.CanParse(firstArg))
            {
                var result = _timeStringParser.Parse(firstArg);
                return new CompileResult(result, DataType.Time);
            }
            var hour = ArgToInt(arguments, 0);
            var min = ArgToInt(arguments, 1);
            var sec = ArgToInt(arguments, 2);

            ThrowArgumentExceptionIf(() => sec < 0 || sec > 59, "Invalid second: " + sec);
            ThrowArgumentExceptionIf(() => min < 0 || min > 59, "Invalid minute: " + min);
            ThrowArgumentExceptionIf(() => min < 0 || hour > 23, "Invalid hour: " + hour);


            var secondsOfThisTime = (double)(hour * 60 * 60 + min * 60 + sec);
            return new CompileResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
        }
    }
}
