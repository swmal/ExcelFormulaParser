using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelFormulaParser.Engine.ExpressionGraph
{
    public class FunctionArgumentExpression : GroupExpression
    {
        public override bool ParentIsLookupFunction
        {
            get
            {
                return base.ParentIsLookupFunction;
            }
            set
            {
                base.ParentIsLookupFunction = value;
                foreach (var child in Children)
                {
                    child.ParentIsLookupFunction = value;
                }
            }
        }
    }
}
