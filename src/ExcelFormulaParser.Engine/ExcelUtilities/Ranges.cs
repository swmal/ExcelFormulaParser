using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Engine.ExcelUtilities
{
    public class Ranges
    {
        private readonly List<RangeAddress> _rangeAddresses = new List<RangeAddress>();

        public virtual void Add(string rangeAddress)
        {
            Add(RangeAddress.Parse(rangeAddress));
        }

        public virtual void Add(RangeAddress rangeAddress)
        {
            _rangeAddresses.Add(rangeAddress);
        }

        public virtual void Clear()
        {
            _rangeAddresses.Clear();
        }

        public virtual void CheckCircularReference()
        {
            _rangeAddresses.Aggregate((x, y) =>
                {
                    if (x.CollidesWith(y))
                    {
                        throw new CircularReferenceException("Circular reference detected");
                    }
                    return x;
                });
        }
    }
}
