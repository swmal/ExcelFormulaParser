﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ExcelFormulaParser.Engine;

namespace ExcelFormulaParser.Tests
{
    [TestClass]
    public class ParsingContextTests
    {
        [TestMethod]
        public void ConfigurationShouldBeSetByFactoryMethod()
        {
            var context = ParsingContext.Create();
            Assert.IsNotNull(context.Configuration);
        }
    }
}