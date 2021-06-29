﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System;
using MsiReader;





namespace UnitTests
{
    [TestClass]
    public class ReaderTest
    {
        [TestMethod]
        public void FileIsLoaded()
        {
            List<String> allFileNames = new List<String>();


            Assert.AreEqual(0,MsiPull.DrawFromMsi("Setupdistex.msi", ref allFileNames));

            Assert.AreEqual("_Validation", allFileNames.First());
            //Assert.IsTrue(someexpressionorboolean);        
        }
        [TestMethod]
        public void FirstItemIsCorrect()
        {
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi("Setupdistex.msi", ref allFileNames);

            Assert.AreEqual("_Validation", allFileNames.First());     
        }
    }
}
