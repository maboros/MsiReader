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
        string projectdir;
        string fullPath;
        [SetUp]
        public void Setup()
        {
            projectdir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            fullPath= projectdir + "\\" + "appData" + "\\68b3ac.msi";
        }

        [Test]
        public void DrawFromMsiReturns0()
        {
            List<String> allFileNames = new List<String>();
            Assert.AreEqual(0, Reader.DrawFromMsi(fullPath, ref allFileNames));
        }

        [Test]
        public void FirstItemInDrawFromMsiIs_Validation()
        {
            List<String> allFileNames = new List<String>();
            Reader.DrawFromMsi(fullPath, ref allFileNames);
            Assert.AreEqual("_Validation", allFileNames.First());
        }
        [Test]
        public void GetItemDataReturns0()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Assert.AreEqual(0, Reader.GetItemData(fullPath, "_Validation", ref columnString, ref columnCount, ref dataString));
        }
        [Test]
        public void FirstItemInGetItemDataReturn10ColumnsFor_Validation()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Reader.GetItemData(fullPath, "_Validation", ref columnString, ref columnCount, ref dataString);
            Assert.AreEqual(10,columnCount);
        }
        [Test]
        public void FirstItemInGetItemDataReturnsTableForFirstColumnFor_Validation()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Reader.GetItemData(fullPath, "_Validation", ref columnString, ref columnCount, ref dataString);
            Assert.AreEqual("Table",columnString.First());
        }
        [Test]
        public void FirstItemInGetItemDataReturnsString()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Reader.GetItemData(fullPath, "_Validation", ref columnString, ref columnCount, ref dataString);
            Assert.AreEqual("_Validation", dataString.First());
        }
        [Test]
        public void FirstItemInGetItemDataReturnsInteger()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Reader.GetItemData(fullPath, "Media", ref columnString, ref columnCount, ref dataString);
            Assert.AreEqual("1", dataString.First());
        }
        [Test]
        public void FirstItemInGetItemDataReturnsBinaryData()
        {
            List<String> columnString = new List<String>();
            List<String> dataString = new List<String>();
            int columnCount = 0;
            Reader.GetItemData(fullPath, "MsiDigitalCertificate", ref columnString, ref columnCount, ref dataString);
            Assert.AreEqual("[Binary data]", dataString.Last());
        }
        [Test]
        public void GetSummaryInformationPullsFirstItem()
        {
            var list = Reader.GetSummaryInformation(fullPath);
            Assert.AreEqual("CodePageString", list.First().name);
            Assert.AreEqual("1252", list.First().value);
        }
    }
}