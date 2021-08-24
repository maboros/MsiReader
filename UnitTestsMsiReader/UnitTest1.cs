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
        string solutiondir;
        [SetUp]
        public void Setup()
        {
            solutiondir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
        }

        [Test]
        public void FileIsLoadedAndDrawFromMsiReturns0()
        {
            List<String> allFileNames = new List<String>();
            Assert.AreEqual(0, MsiPull.DrawFromMsi(solutiondir + "\\" + "appData" + "\\68b3ac.msi", ref allFileNames));
        }

        [Test]
        public void FirstItemInDrawFromMsiIsCorrect()
        {
            List<String> allFileNames = new List<String>();
            MsiPull.DrawFromMsi(solutiondir + "\\" + "appData" + "\\68b3ac.msi", ref allFileNames);
            Assert.AreEqual("_Validation", allFileNames.First());
        }

        [Test]
        public void GetSummaryInformationPullsFirstItem()
        {
            var list = MsiPull.getSummaryInformation(solutiondir + "\\" + "appData" + "\\68b3ac.msi");
            Assert.AreEqual("CodePageString", list.First().name);
            Assert.AreEqual("1252", list.First().value);
        }
    }
}