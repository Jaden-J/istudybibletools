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

            var result = TestsHelper.AnalyzeString(ref _oneNoteApp, "TestScenario1", input);

            Assert.AreEqual(expected, StringUtils.GetText(result.OutputHTML));
            Assert.IsTrue(result.FoundVerses.Count == 10);
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 1:16")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:13")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:14")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:15")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:16")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:17")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 18")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 19")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 11:1")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 11:2")));            
        }

        [TestMethod]
        public void TestScenario2()
        {
            var input = "тест Лк 1:16, 10:13-17,18-19; 11:1-2 тест";

            var result = TestsHelper.AnalyzeString(ref _oneNoteApp, "TestScenario2", input);

            Assert.AreEqual(StringUtils.GetText(input), StringUtils.GetText(result.OutputHTML));
            Assert.IsTrue(result.FoundVerses.Count == 10);
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 1:16")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:13")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:14")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:15")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:16")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:17")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:18")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 10:19")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 11:1")));
            Assert.IsTrue(result.FoundVerses.Contains(new VersePointer("Лк 11:2")));
        }
    }
}
