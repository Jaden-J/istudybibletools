using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BibleCommon.Helpers;
using System.Xml.Linq;
using BibleCommon.Common;

namespace Tests
{
    [TestClass]
    public class LinkParserTest
    {
        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;      
    
        [TestInitialize]
        public void Init()
        {
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();          
        }

        [TestCleanup]
        public void Done()
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);            
        }

        [TestMethod]
        public void TestScenario1()
        {
            var input = "тест Лк 1:16, 10:13-17;18-19; 11:1-2 тест";
            var expected = "тест Лк 1:16, 10:13-17; 18-19; 11:1-2 тест";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario1", input);

            Assert.AreEqual(expected, StringUtils.GetText(result.OutputHTML));
            TestHelper.CheckVerses(result, "Лк 1:16", "Лк 10:13", "Лк 10:14", "Лк 10:15", "Лк 10:16", 
                "Лк 10:17", "Лк 18", "Лк 19", "Лк 11:1", "Лк 11:2");            
        }

        [TestMethod]
        public void TestScenario2()
        {
            var input = "тест Лк 1:16, 10:13-17,18-19; 11:1-2 тест"; 

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario2", input);

            Assert.AreEqual(StringUtils.GetText(input), StringUtils.GetText(result.OutputHTML));
            TestHelper.CheckVerses(result, "Лк 1:16", "Лк 10:13", "Лк 10:14", 
                "Лк 10:15", "Лк 10:16", "Лк 10:17", "Лк 10:18", "Лк 10:19", "Лк 11:1", "Лк 11:2");  
        }

        [TestMethod]
        public void TestScenario3()
        {
            var input = "Этот тест из 1 Ин 1 был подготвлен в (:2) и :3-4 и в :7-6, _:8_ стихах. А в 2-е Ин 1:3-5,6 тоже интересная инфа о {:7}. И о 2Тим 1:1,2-3";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario3", input);

            Assert.AreEqual(StringUtils.GetText(input), StringUtils.GetText(result.OutputHTML));
            TestHelper.CheckVerses(result, "1Ин 1", "1Ин 1:2", "1Ин 1:3", "1Ин 1:4", "1Ин 1:7", "1Ин 1:8", 
                "2Ин 1:3", "2Ин 1:4", "2Ин 1:5", "2Ин 1:6", "2Ин 1:7", "1Тим 1:1", "1Тим 1:2", "1Тим 1:3");       
        }

        [TestMethod]
        public void TestScenario4()
        {
            var input = "1 Лк 1:1, 2";
            var expected = "1 Лк 1:1,2";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario4", input);

            Assert.AreEqual(expected, StringUtils.GetText(result.OutputHTML));
            TestHelper.CheckVerses(result, "Лк 1:1", "Лк 1:2");
        }
    }
}
