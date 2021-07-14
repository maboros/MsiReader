using MsiReader;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace UnitTestsMsiReader
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void FileIsLoaded()
        {
            List<String> allFileNames = new List<String>();
            string solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            Assert.AreEqual(0, MsiPull.DrawFromMsi(solutiondir + "\\" + "appData" + "\\68b3ac.msi", ref allFileNames));
        }

        [Test]
        public void FirstItemIsCorrect()
        {
            List<String> allFileNames = new List<String>();
            string solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            Assert.AreEqual(0, MsiPull.DrawFromMsi(solutiondir + "\\" + "appData" + "\\68b3ac.msi", ref allFileNames));
            Assert.AreEqual("_Validation", allFileNames.First());
        }

        [Test]
        public void getSummaryInformationPullsFirstItem()
        {
            string solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            var list = MsiPull.getSummaryInformation(solutiondir + "\\" + "appData" + "\\68b3ac.msi");
            Assert.AreEqual("CodePageString", list.First().name);
            Assert.AreEqual("1252", list.First().value);
        }
    }
}