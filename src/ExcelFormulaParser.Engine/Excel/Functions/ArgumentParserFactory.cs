using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.ExpressionGraph;

namespace ExcelFormulaParser.Engine.Excel.Functions
{
    public class ArgumentParserFactory
    {
        public virtual ArgumentParser CreateArgumentParser(DataType dataType)
        {
            switch (dataType)
            {
                case DataType.Integer:
                    return new IntArgumentParser();
                case DataType.Boolean:
                    return new BoolArgumentParser();
                case DataType.Decimal:
                    return new DoubleArgumentParser();
                default:
                    throw new InvalidOperationException("non supported argument parser type " + dataType.ToString());
            }
        }
    }
}
