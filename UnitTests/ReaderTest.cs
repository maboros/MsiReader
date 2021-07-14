using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using System;
using MsiReader;
using System.IO;

namespace UnitTests
{
    [TestClass]
    public class ReaderTest
    {
        [TestMethod]
        public void FileIsLoaded()
        {
            List<String> allFileNames = new List<String>();
            string solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            Assert.AreEqual(0, MsiPull.DrawFromMsi("Setupdistex.msi", ref allFileNames));
            Assert.AreEqual("_Validation", allFileNames.First());      
        }
        [TestMethod]
        public void FirstItemIsCorrect()
        {
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi("Setupdistex.msi", ref allFileNames);

            Assert.AreEqual("_Validation", allFileNames.First());     
        }

        [TestMethod]
        public void getSummaryInformationPullsFirstItem()
        {
            string solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            var list=MsiPull.getSummaryInformation(solutiondir + "\\" + "appData" + "\\68b3ac.msi");
            Assert.AreEqual("CodePageString", list.First().name);
            Assert.AreEqual("1252", list.First().value);
        }
    }

}
