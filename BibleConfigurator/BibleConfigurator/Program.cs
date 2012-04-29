using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using BibleCommon.Common;
using System.Xml.Serialization;
using System.IO;

namespace BibleConfigurator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(params string[] args)
        {
            GenerateModuleInfo();


            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm(args));
        }

        private static void GenerateModuleInfo()
        {
            ModuleInfo module = new ModuleInfo()
            {
                Version = "1.0",
                Name = "Синодальный перевод (Русский язык)",
                Notebooks = new List<NotebookInfo>() 
                {
                    new NotebookInfo() { Type = NotebookType.Single, Name = "Holy Bible.onepkg", SectionGroups = new List<BibleCommon.Common.SectionGroupInfo>()
                    {
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.Bible, Name="Библия" },
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleStudy, Name="Изучение Библии" },
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleComments, Name="Комментарии к Библии" },
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleNotesPages, Name="Сводные заметок" }
                    } },
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Библия.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Изучение Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Комментарии к Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Сводные заметок.onepkg" }
                },
                BibleStructure = new BibleStructureInfo()
                {
                    BibleBooks = new List<BibleBookInfo>()
                    {
                       new BibleBookInfo() { Name = "Бытие", SectionName = "01. Бытие", Shortenings = new List<string>() { "быт", "бт", "бытие" } },
                       new BibleBookInfo() { Name = "Исход", SectionName = "02. Исход", Shortenings = new List<string>() { "исх", "исход" } },
                       new BibleBookInfo() { Name = "Левит", SectionName = "03. Левит", Shortenings = new List<string>() { "лев", "лв", "левит" } },
                       new BibleBookInfo() { SectionName = "04. Числа", new StringCollection() { "чис", "чс", "числ", "числа" } },
                       new BibleBookInfo() { SectionName = "05. Второзаконие", new StringCollection() { "втор", "вт", "втрзк", "второзаконие" } },
                       new BibleBookInfo()  { "06. Иисус Навин", new StringCollection() { "иис.нав", "нав", "иисус навин", "ииснав", "ис. нав", "ис.нав", "навин" } },
                        new BibleBookInfo() { "07. Судьи", new StringCollection() { "суд", "сд", "судьи", "судей" } },
                        new BibleBookInfo() { "08. Руфь", new StringCollection() { "руф", "рф", "руфь" } },
                        new BibleBookInfo() { "09. 1-я Царств", new StringCollection() { "1цар", "1 цар", "1цр", "1ц", "1царств", "1-я царств", "1 царств" } },
                        new BibleBookInfo() { "10. 2-я Царств", new StringCollection() { "2цар", "2 цар", "2цр", "2ц", "2царств", "2-я царств", "2 царств" } },
                        new BibleBookInfo() { "11. 3-я Царств", new StringCollection() { "3цар", "3 цар", "3цр", "3ц", "3царств", "3-я царств", "3 царств" } },
                        new BibleBookInfo() { "12. 4-я Царств", new StringCollection() { "4цар", "4 цар", "4цр", "4ц", "4царств", "4-я царств", "4 царств" } },
                        new BibleBookInfo() { "13. 1-я Паралипоменон", new StringCollection() { "1пар", "1пр", "1 пар", "1-я паралипоменон" } },
                        new BibleBookInfo() { "14. 2-я Паралипоменон", new StringCollection() { "2пар", "2пр", "2 пар", "2-я паралипоменон" } },
                        new BibleBookInfo() { "15. Ездра", new StringCollection() { "ездр", "езд", "ез", "ездра" } },
                        new BibleBookInfo() { "16. Неемия", new StringCollection() { "неем", "нм", "неемия" } },
                        new BibleBookInfo() { "17. Есфирь", new StringCollection() { "есф", "ес", "есфирь" } },
                        new BibleBookInfo() { "18. Иов", new StringCollection() { "иов", "ив" } },
                        new BibleBookInfo() { "19. Псалтирь", new StringCollection() { "пс", "псалт", "псал", "псл", "псалом", "псалтырь", "псалмы", "псалтирь" } },
                        new BibleBookInfo() { "20. Притчи", new StringCollection() { "прит", "притч", "при", "притчи", "притча", "пр" } },
                        new BibleBookInfo() { "21. Екклесиаст", new StringCollection() { "еккл", "ек", "екк", "екклесиаст" } },
                        new BibleBookInfo() { "22. Песня Песней", new StringCollection() { "песн", "пес", "псн", "песн.песней", "песни", "песня песней" } },
                        new BibleBookInfo() { "23. Исаия", new StringCollection() { "ис", "иса", "исаия" } },
                        new BibleBookInfo() { "24. Иеремия", new StringCollection() { "иер", "иерем", "иеремия" } },
                        new BibleBookInfo() { "25. Плач Иеремии", new StringCollection() { "плач", "плч", "пл", "пл.иер", "плач иеремии" } },
                        new BibleBookInfo() { "26. Иезекииль", new StringCollection() { "иез", "иезек", "иезекииль" } },
                        new BibleBookInfo() { "27. Даниил", new StringCollection() { "дан", "дн", "днл", "даниил" } },
                        new BibleBookInfo() { "28. Осия", new StringCollection() { "ос", "осия" } },
                        new BibleBookInfo() { "29. Иоиль", new StringCollection() { "иоил", "ил", "иоиль" } },
                        new BibleBookInfo() { "30. Амос", new StringCollection() { "ам", "амс", "амос" } },
                        new BibleBookInfo() { "31. Авдий", new StringCollection() { "авд", "авдий" } },
                        new BibleBookInfo() { "32. Иона", new StringCollection() { "ион", "иона" } },
                        new BibleBookInfo() { "33. Михей", new StringCollection() { "мих", "мх", "михей", "михея" } },
                        new BibleBookInfo() { "34. Наум", new StringCollection() { "наум" } },
                        new BibleBookInfo() { "35. Аввакум", new StringCollection() { "авв", "аввак", "аввакум" } },
                        new BibleBookInfo() { "36. Софония", new StringCollection() { "соф", "софон", "софония" } },
                        new BibleBookInfo() { "37. Аггей", new StringCollection() { "агг", "аггей" } },
                        new BibleBookInfo() { "38. Захария", new StringCollection() { "захария", "зах", "зхр", "захар" } },
                        new BibleBookInfo() { "39. Малахия", new StringCollection() { "мал", "малах", "млх", "малахия" } },
                        new BibleBookInfo() { "01. От Матфея", new StringCollection() { "мат", "матф", "мтф", "мф", "мт", "матфей", "матфея", "от матфея" } },
                        new BibleBookInfo() { "02. От Марка", new StringCollection() { "мар", "марк", "мрк", "мр", "марка", "мк", "от марка" } },
                        new BibleBookInfo() { "03. От Луки", new StringCollection() { "лук", "лк", "лука", "луки", "от луки" } },
                        new BibleBookInfo() { "04. От Иоанна", new StringCollection() { "иоан", "ин", "иоанн", "иоанна", "от иоанна" } },
                        new BibleBookInfo() { "05. Деяния", new StringCollection() { "деян", "дея", "д.а", "деяния", "деяния апостолов" } },
                        new BibleBookInfo() { "06. Иакова", new StringCollection() { "иак", "ик", "иаков", "иакова" } },
                        new BibleBookInfo() { "07. 1-е Петра", new StringCollection() { "1пет", "1 пет", "1пт", "1птр", "1 птр", "1петр", "1 петр", "1петра", "1-е петра", "1 петра", "1 петр", "1-е петр" } },
                        new BibleBookInfo() { "08. 2-е Петра", new StringCollection() { "2пет", "2 пет", "2пт", "2птр", "2 птр", "2петр", "2 петр", "2петра", "2-е петра", "2 петра", "2 петр", "2-е петр" } },
                        new BibleBookInfo() { "09. 1-е Иоанна", new StringCollection() { "1иоан", "1ин", "1иоанн", "1иоанна", "1-е иоанна", "1 иоанн", "1 иоанна", "1 ин", "1 иоан", "1-е иоанн", "1-е иоан", "1-е ин" } },
                        new BibleBookInfo() { "10. 2-е Иоанна", new StringCollection() { "2иоан", "2ин", "2иоанн", "2иоанна", "2-е иоанна", "2 иоанн", "2 иоанна", "2 ин", "2 иоан", "2-е иоанн", "2-е иоан", "2-е ин" } },
                        new BibleBookInfo() { "11. 3-е Иоанна", new StringCollection() { "3иоан", "3ин", "3иоанн", "3иоанна", "3-е иоанна", "3 иоанн", "3 иоанна", "3 ин", "3 иоан", "3-е иоанн", "3-е иоан", "3-е ин" } },
                        new BibleBookInfo() { "12. Иуды", new StringCollection() { "иуд", "ид", "иуда", "иуды" } },
                        new BibleBookInfo() { "13. К Римлянам", new StringCollection() { "рим", "римл", "римлянам", "к римлянам" } },
                        new BibleBookInfo() { "14. 1-е Коринфянам", new StringCollection() { "1кор", "1 кор", "1коринф", "1коринфянам", "1-е коринфянам", "1 коринфянам", "1 коринф", "1-е коринф", "1-е кор" } },
                        new BibleBookInfo() { "15. 2-е Коринфянам", new StringCollection() { "2кор", "2 кор", "2коринф", "2коринфянам", "2-е коринфянам", "2 коринфянам", "2 коринф", "2-е коринф", "2-е кор" } },
                        new BibleBookInfo() { "16. К Галатам", new StringCollection() { "гал", "галат", "галатам", "к галатам" } },
                        new BibleBookInfo() { "17. К Ефесянам", new StringCollection() { "еф", "ефес", "ефесянам", "к ефесянам" } },
                        new BibleBookInfo() { "18. К Филиппийцам", new StringCollection() { "фил", "флп", "филип", "филиппийцам", "к филиппийцам" } },
                        new BibleBookInfo() { "19. К Колоссянам", new StringCollection() { "кол", "колос", "колоссянам", "к колоссянам" } },
                        new BibleBookInfo() { "20. 1-е Фессалоникийцам", new StringCollection() { "1фесс", "1фес", "1фессалоникийцам", "1 фес", "1 фесс", "1-е фессалоникийцам", "1 фессалоникийцам", "1-е фесс", "1-е фес" } },
                        new BibleBookInfo() { "21. 2-е Фессалоникийцам", new StringCollection() { "2фесс", "2фес", "2фессалоникийцам", "2 фес", "2 фесс", "2-е фессалоникийцам", "2 фессалоникийцам", "2-е фесс", "2-е фес" } },
                        new BibleBookInfo() { "22. 1-е Тимофею", new StringCollection() { "1тим", "1тимоф", "1тимофею", "1 тимофею", "1-е тимофею", "1 тим", "1 тимоф", "1-е тимоф", "1-е тим" } },
                        new BibleBookInfo() { "23. 2-е Тимофею", new StringCollection() { "2тим", "2тимоф", "2тимофею", "2 тимофею", "2-е тимофею", "2 тим", "2 тимоф", "2-е тимоф", "2-е тим" } },
                        new BibleBookInfo() { "24. К Титу", new StringCollection() { "тит", "титу", "к титу" } },
                        new BibleBookInfo() { "25. К Филимону", new StringCollection() { "флм", "филимон", "филимону", "к филимону" } },
                        new BibleBookInfo() { "26. К Евреям", new StringCollection() { "евр", "евреям", "к евреям" } },
                        new BibleBookInfo() { "27. Откровение", new StringCollection() { "откр", "отк", "откровен", "апок", "откровение", "апокалипсис" } }
                    }
                }
            };

            XmlSerializer ser = new XmlSerializer(typeof(ModuleInfo));
            using (var fs = new FileStream("c:\\moduleInfo.xml", FileMode.Create))
            {

                ser.Serialize(fs, module);
                fs.Flush();

            }
        }
    }
}
