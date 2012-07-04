using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;

namespace ExcelFormulaParser.Tests.ExcelUtilities
{
    [TestClass]
    public class RangeAddressTests
    {
        [TestInitialize]
        public void Setup()
        {

        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ParseShouldThrowIfSuppliedAddressIsNull()
        {
            RangeAddress.Parse(null);
        }

        [TestMethod]
        public void ParseShouldReturnAndInstanceWithColPropertiesSet()
        {
            var address = RangeAddress.Parse("A2");
            Assert.AreEqual(0, address.FromCol, "FromCol was not 0");
            Assert.AreEqual(0, address.ToCol, "ToCol was not 0");
        }

        [TestMethod]
        public void ParseShouldReturnAndInstanceWithRowPropertiesSet()
        {
            var address = RangeAddress.Parse("A2");
            Assert.AreEqual(1, address.FromRow, "FromRow was not 1");
            Assert.AreEqual(1, address.ToRow, "ToRow was not 1");
        }

        [TestMethod]
        public void ParseShouldReturnAnInstanceWithFromAndToColSetWhenARangeAddressIsSupplied()
        {
            var address = RangeAddress.Parse("A1:B2");
            Assert.AreEqual(0, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void ParseShouldReturnAnInstanceWithFromAndToRowSetWhenARangeAddressIsSupplied()
        {
            var address = RangeAddress.Parse("A1:B3");
            Assert.AreEqual(0, address.FromRow);
            Assert.AreEqual(2, address.ToRow);
        }

        [TestMethod]
        public void ParseShouldSetWorksheetNameIfSuppliedInAddress()
        {
            var address = RangeAddress.Parse("Ws!A1");
            Assert.AreEqual("Ws", address.Worksheet);
        }

        [TestMethod]
        public void CollideShouldReturnTrueIfRangesCollides()
        {
            var address1 = RangeAddress.Parse("A1:A6");
            var address2 = RangeAddress.Parse("A5");
            Assert.IsTrue(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesDoesNotCollide()
        {
            var address1 = RangeAddress.Parse("A1:A6");
            var address2 = RangeAddress.Parse("A8");
            Assert.IsFalse(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CollideShouldReturnFalseIfRangesCollidesButWorksheetNameDiffers()
        {
            var address1 = RangeAddress.Parse("Ws!A1:A6");
            var address2 = RangeAddress.Parse("A5");
            Assert.IsFalse(address1.CollidesWith(address2));
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithStringAddressSet()
        {
            var address = RangeAddress.Create(0, 0);
            Assert.AreEqual("A1", address.ToString());
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToColSet()
        {
            var address = RangeAddress.Create(1, 0);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToRowSet()
        {
            var address = RangeAddress.Create(0, 1);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(1, address.ToRow);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithWorksheetSetToEmptyString()
        {
            var address = RangeAddress.Create(0, 1);
            Assert.AreEqual(string.Empty, address.Worksheet);
        }
    }
}
