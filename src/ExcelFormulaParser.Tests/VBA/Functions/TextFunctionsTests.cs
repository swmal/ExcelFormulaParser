using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.VBA.Functions;
using ExcelFormulaParser.Engine.ExpressionGraph;
using ExcelFormulaParser.Engine.VBA.Functions.Text;

namespace ExcelFormulaParser.Tests.VBA.Functions.Text
{
    [TestClass]
    public class TextFunctionsTests
    {
        [TestMethod]
        public void CStrShouldConvertNumberToString()
        {
            var func = new CStr();
            var result = func.Execute(new object[] { 1 });
            Assert.AreEqual(DataType.String, result.DataType);
            Assert.AreEqual("1", result.Result);
        }

        [TestMethod]
        public void LenShouldReturnStringsLength()
        {
            var func = new Len();
            var result = func.Execute(new object[] { "abc" });
            Assert.AreEqual(3, result.Result);
        }

        [TestMethod]
        public void LowerShouldReturnLowerCaseString()
        {
            var func = new Lower();
            var result = func.Execute(new object[] { "ABC" });
            Assert.AreEqual("abc", result.Result);
        }

        [TestMethod]
        public void UpperShouldReturnUpperCaseString()
        {
            var func = new Upper();
            var result = func.Execute(new object[] { "abc" });
            Assert.AreEqual("ABC", result.Result);
        }

        [TestMethod]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var func = new Left();
            var result = func.Execute(new object[] { "abcd", 2 });
            Assert.AreEqual("ab", result.Result);
        }

        [TestMethod]
        public void RightShouldReturnSubstringFromRight()
        {
            var func = new Right();
            var result = func.Execute(new object[] { "abcd", 2 });
            Assert.AreEqual("cd", result.Result);
        }

        [TestMethod]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var func = new Mid();
            var result = func.Execute(new object[] { "abcd", 1, 2 });
            Assert.AreEqual("bc", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
        {
            var func = new Replace();
            var result = func.Execute(new object[] { "testar", 1, 2, "hej" });
            Assert.AreEqual("hejstar", result.Result);
        }

        [TestMethod]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
        {
            var func = new Replace();
            var result = func.Execute(new object[] { "testar", 3, 3, "hej" });
            Assert.AreEqual("tehejr", result.Result);
        }

        [TestMethod]
        public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
        {
            var func = new Substitute();
            var result = func.Execute(new object[] { "testar testar", "es", "xx" });
            Assert.AreEqual("txxtar txxtar", result.Result);
        }

        [TestMethod]
        public void ConcatenateShouldConcatenateThreeStrings()
        {
            var func = new Concatenate();
            var result = func.Execute(new object[] { "One", "Two", "Three" });
            Assert.AreEqual("OneTwoThree", result.Result);
        }
    }
}
