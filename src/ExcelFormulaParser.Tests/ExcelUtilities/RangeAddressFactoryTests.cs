﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine.ExcelUtilities;
using Rhino.Mocks;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests.ExcelUtilities
{
    [TestClass]
    public class RangeAddressFactoryTests
    {
        private RangeAddressFactory _factory;

        [TestInitialize]
        public void Setup()
        {
            var provider = MockRepository.GenerateStub<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void CreateShouldThrowIfSuppliedAddressIsNull()
        {
            _factory.Create(null);
        }

        [TestMethod]
        public void CreateShouldReturnAndInstanceWithColPropertiesSet()
        {
            var address = _factory.Create("A2");
            Assert.AreEqual(0, address.FromCol, "FromCol was not 0");
            Assert.AreEqual(0, address.ToCol, "ToCol was not 0");
        }

        [TestMethod]
        public void CreateShouldReturnAndInstanceWithRowPropertiesSet()
        {
            var address = _factory.Create("A2");
            Assert.AreEqual(1, address.FromRow, "FromRow was not 1");
            Assert.AreEqual(1, address.ToRow, "ToRow was not 1");
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToColSetWhenARangeAddressIsSupplied()
        {
            var address = _factory.Create("A1:B2");
            Assert.AreEqual(0, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToRowSetWhenARangeAddressIsSupplied()
        {
            var address = _factory.Create("A1:B3");
            Assert.AreEqual(0, address.FromRow);
            Assert.AreEqual(2, address.ToRow);
        }

        [TestMethod]
        public void CreateShouldSetWorksheetNameIfSuppliedInAddress()
        {
            var address = _factory.Create("Ws!A1");
            Assert.AreEqual("Ws", address.Worksheet);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithStringAddressSet()
        {
            var address = _factory.Create(0, 0);
            Assert.AreEqual("A1", address.ToString());
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToColSet()
        {
            var address = _factory.Create(1, 0);
            Assert.AreEqual(1, address.FromCol);
            Assert.AreEqual(1, address.ToCol);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithFromAndToRowSet()
        {
            var address = _factory.Create(0, 1);
            Assert.AreEqual(1, address.FromRow);
            Assert.AreEqual(1, address.ToRow);
        }

        [TestMethod]
        public void CreateShouldReturnAnInstanceWithWorksheetSetToEmptyString()
        {
            var address = _factory.Create(0, 1);
            Assert.AreEqual(string.Empty, address.Worksheet);
        }
    }
}