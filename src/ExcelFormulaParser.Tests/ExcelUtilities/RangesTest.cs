using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;
using ExcelFormulaParser.Engine.Exceptions;

namespace ExcelFormulaParser.Tests.ExcelUtilities
{
    [TestClass]
    public class RangesTest
    {
        //[TestMethod]
        //public void ShouldValidateWithoutExceptionIfRangesDoesNotCollide()
        //{
        //    var ranges = new Ranges();
        //    ranges.Add("A1");
        //    ranges.CheckCircularReference(RangeAddress.Parse("A2"));
        //}

        //[TestMethod, ExpectedException(typeof(CircularReferenceException))]
        //public void ShouldThrowExceptionIfRangesCollide()
        //{
        //    var ranges = new Ranges();
        //    ranges.Add("A1:A3");
        //    ranges.CheckCircularReference(RangeAddress.Parse("A2"));
        //}
    }
}
