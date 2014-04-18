using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using BibleCommon.Helpers;
using System.Xml.Linq;
using BibleCommon.Common;
using BibleCommon.Services;

namespace Tests
{
    [TestClass]
    public class LinkParserTests
    {
        private static Microsoft.Office.Interop.OneNote.Application _oneNoteApp;      
    
        [TestInitialize]
        public void Init()
        {
            _oneNoteApp = OneNoteUtils.CreateOneNoteAppSafe();
            SettingsManager.Instance.UseCommaDelimeter = true;
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
            TestHelper.CheckVerses(expected, result, "Лк 1:16", "Лк 10:13", "Лк 10:14", "Лк 10:15", "Лк 10:16", 
                "Лк 10:17", "Лк 18", "Лк 19", "Лк 11:1", "Лк 11:2");            
        }

        [TestMethod]
        public void TestScenario2()
        {
            var input = "тест Лк 1:16, 10:13-17,18-19; 11:1-2 тест"; 

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario2", input);            
            TestHelper.CheckVerses(input, result, "Лк 1:16", "Лк 10:13", "Лк 10:14", 
                "Лк 10:15", "Лк 10:16", "Лк 10:17", "Лк 10:18", "Лк 10:19", "Лк 11:1", "Лк 11:2");  
        }

        [TestMethod]
        public void TestScenario3()
        {
            var input = "Этот тест из 1 Ин 1 был подготвлен в (:2) и :3-4 и в :7-6, _:8_ стихах. А в 2-е Ин 1:3-5,6 тоже интересная инфа о {:7}. И о 2Тим 1:1,2-3";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario3", input);            
            TestHelper.CheckVerses(input, result, "1Ин 1", "1Ин 1:2", "1Ин 1:3", "1Ин 1:4", "1Ин 1:7", "1Ин 1:8", 
                "2Ин 1:3", "2Ин 1:4", "2Ин 1:5", "2Ин 1:6", "2Ин 1:7", "2Тим 1:1", "2Тим 1:2", "2Тим 1:3");       
        }

        [TestMethod]
        public void TestScenario4()
        {
            var input = "1 Лк 1:1, 2";
            var expected = "1 Лк 1:1,2";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario4", input);            
            TestHelper.CheckVerses(expected, result, "Лк 1:1", "Лк 1:2");
        }

        [TestMethod]
        public void TestScenario5()
        {
            var input = "Ин 1: вот и Отк 5(синодальный перевод) и Деяния 1:5,6: вот";            

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario5", input);            
            TestHelper.CheckVerses(input, result, "Ин 1", "Отк 5", "Деян 1:5", "Деян 1:6");
        }

        [TestMethod]
        public void TestScenario6()
        {
            var input = "Ин 1:50-2:2,3-4";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario6", input);            
            TestHelper.CheckVerses(input, result, "Ин 1:50", "Ин 1:51", "Ин 2:1", "Ин 2:2", "Ин 2:3", "Ин 2:4");
        }

        [TestMethod]
        public void TestScenario7()
        {
            var input = "Ин 1:50-2:2,4-5";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario7", input);
            TestHelper.CheckVerses(input, result, "Ин 1:50", "Ин 1:51", "Ин 2:1", "Ин 2:2", "Ин 2:4", "Ин 2:5");
        }

        [TestMethod]
        public void TestScenario8()
        {
            var input = ":1-2 как и в :3,4-5";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario8 (1Кор 1)", input);
            TestHelper.CheckVerses(input, result, "1Кор 1", "1Кор 1:1", "1Кор 1:2", "1Кор 1:3", "1Кор 1:4", "1Кор 1:5");
        }

        [TestMethod]
        public void TestScenario9()
        {
            var input = "Ps 89:1-2";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario9", input);
            TestHelper.CheckVerses(input, result, "Пс 88:1", "Пс 88:2", "Пс 88:3");
        }

        [TestMethod]
        public void TestScenario10()
        {
            var input = "В Ин 1,1 написано. И в 1,2 веке про это писали! Про :2 - тоже";
            var expected = "В Ин 1:1 написано. И в 1,2 веке про это писали! Про :2 - тоже";
            
            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario10", input);
            TestHelper.CheckVerses(expected, result, "Ин 1:1", "Ин 1:2");
        }

        [TestMethod]
        public void TestScenario11()
        {
            var input = "в 1 Ин 1,2-3 и в Иисуса Навина 2-3 было написано про 1-е Кор 1,2-3,4-5;6-7,8-9,10 и в :7";
            var expected = "в 1 Ин 1:2-3 и в Иисуса Навина 2-3 было написано про 1-е Кор 1:2-3,4-5; 6-7, 8-9, 10 и в :7";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario11", input);
            TestHelper.CheckVerses(expected, result, "1 Ин 1:2", "1 Ин 1:3", "Нав 2", "Нав 3", 
                "1Кор 1:2", "1Кор 1:3", "1Кор 1:4", "1Кор 1:5", "1Кор 6", "1Кор 7", "1Кор 8", "1Кор 9", "1Кор 10", "1Кор 10:7");
        }

        [TestMethod]
        public void TestScenario12()
        {
            var input = "Ин 1,2,3 и ещё: Марка 1,2, 3: а потом Лк 1,2- 3";
            var expected = "Ин 1:2,3 и ещё: Марка 1:2,3: а потом Лк 1:2-3";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario12", input);
            TestHelper.CheckVerses(expected, result, "Ин 1:2", "Ин 1:3", "Мк 1:2", "Мк 1:3", "Лк 1:2", "Лк 1:3");
        }

        [TestMethod]
        public void TestScenario13()
        {
            var input = "<span lang=en>1</span><span lang=ru>И</span><span lang=ru>н</span><span lang=ru> </span><span lang=ru>1</span><span lang=ru>:</span><span lang=ru>1</span> и <span lang=ru>:</span><span lang=ru>7</span>";
            var expected = "1Ин 1:1 и :7";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario13", input);
            TestHelper.CheckVerses(expected, result, "1Ин 1:1", "1Ин 1:7");
        }

        [TestMethod]
        public void TestScenario14()
        {
            var input = "I Cor 6:7, II Tim 2:3";            

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario14", input);
            TestHelper.CheckVerses(input, result, "1Кор 6:7", "2 Тим 2:3");
        }

        [TestMethod]
        public void TestScenario15()
        {
            var input = "<span lang=ru>Исх. 13,1</span><span lang=ro>4</span><span lang=ru>,</span><span lang=se-FI>15</span><span lang=ru>,20.</span>";            
            var expected = "Исх. 13:14,15,20.";

            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario15", input);
            TestHelper.CheckVerses(expected, result, "Исх 13:14", "Исх 13:15", "Исх 13:20");
        }        

        [TestMethod]
        public void TestScenario16()
        {
            var input = "<span lang=ru>Вот Ин 1</span><span lang=en-US>:</span><span lang=ru>12 где в </span><span lang=ro>:</span><span lang=se-FI>13</span>";
            var expected = "Вот Ин 1:12 где в :13";
            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario16", input);
            TestHelper.CheckVerses(expected, result, "Ин 1:12", "Ин 1:13");
        }


        [TestMethod]
        public void TestScenario17()
        {
            var input = "Иуда 14,15";            
            var result = TestHelper.AnalyzeString(ref _oneNoteApp, "TestScenario17", input);
            TestHelper.CheckVerses(input, result, "Иуд 14", "Иуд 15");
        }                
    }
}