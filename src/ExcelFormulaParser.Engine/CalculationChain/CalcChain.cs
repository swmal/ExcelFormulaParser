using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Utilities;

namespace ExcelFormulaParser.Engine.CalculationChain
{
    public class CalcChain
    {
        private CalcChain(IdProvider idProvider)
        {
            Id = idProvider.NewId();
        }

        public static CalcChain Create(IdProvider idProvider)
        {
            Require.That(idProvider).Named("idProvider").IsNotNull();
            return new CalcChain(idProvider);
        }

        public object Id { get; private set; }

        public long Count { get; private set; }

        public CalcCell Current
        {
            get;
            private set;
        }

        public void Add(CalcCell cell)
        {
            Count++;
            Current = cell;
        }
    }
}
