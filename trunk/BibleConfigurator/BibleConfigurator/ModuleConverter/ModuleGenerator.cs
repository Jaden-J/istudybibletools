using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.Xml.Serialization;
using System.IO;

namespace BibleConfigurator.ModuleConverter
{
    public static class ModuleGenerator
    {
        public static void GenerateModuleInfo()
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
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleComments, Name="Комментарии к Библии" }                        
                    } },
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Библия.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Изучение Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Комментарии к Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Сводные заметок.onepkg" }
                },
                BibleStructure = new BibleStructureInfo()
                {
                    OldTestamentName = "Ветхий Завет",
                    NewTestamentName = "Ветхий Завет",
                    Alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя",
                    BibleBooks = new List<BibleBookInfo>()
                    {
                        new BibleBookInfo() { Name = "Бытие", SectionName = "01. Бытие", Shortenings = new List<string>() { "быт", "бт", "бытие" } },
                        new BibleBookInfo() { Name = "Исход", SectionName = "02. Исход", Shortenings = new List<string>() { "исх", "исход" } },
                        new BibleBookInfo() { Name = "Левит", SectionName = "03. Левит", Shortenings = new List<string>() { "лев", "лв", "левит" } },
                        new BibleBookInfo() { Name = "Числа", SectionName = "04. Числа", Shortenings = new List<string>() { "чис", "чс", "числ", "числа" } },
                        new BibleBookInfo() { Name = "Второзаконие", SectionName = "05. Второзаконие", Shortenings = new List<string>() { "втор", "вт", "втрзк", "второзаконие" } },
                        new BibleBookInfo() { Name = "Иисус Навин", SectionName = "06. Иисус Навин", Shortenings = new List<string>() { "иис.нав", "нав", "иисус навин", "ииснав", "ис. нав", "ис.нав", "навин" } },
                        new BibleBookInfo() { Name = "Судьи", SectionName = "07. Судьи", Shortenings = new List<string>() { "суд", "сд", "судьи", "судей" } },
                        new BibleBookInfo() { Name = "Руфь", SectionName = "08. Руфь", Shortenings = new List<string>() { "руф", "рф", "руфь" } },
                        new BibleBookInfo() { Name = "1-я Царств", SectionName = "09. 1-я Царств", Shortenings = new List<string>() { "1цар", "1 цар", "1цр", "1ц", "1царств", "1-я царств", "1 царств" } },
                        new BibleBookInfo() { Name = "2-я Царств", SectionName = "10. 2-я Царств", Shortenings = new List<string>() { "2цар", "2 цар", "2цр", "2ц", "2царств", "2-я царств", "2 царств" } },
                        new BibleBookInfo() { Name = "3-я Царств", SectionName = "11. 3-я Царств", Shortenings = new List<string>() { "3цар", "3 цар", "3цр", "3ц", "3царств", "3-я царств", "3 царств" } },
                        new BibleBookInfo() { Name = "4-я Царств", SectionName = "12. 4-я Царств", Shortenings = new List<string>() { "4цар", "4 цар", "4цр", "4ц", "4царств", "4-я царств", "4 царств" } },
                        new BibleBookInfo() { Name = "1-я Паралипоменон", SectionName = "13. 1-я Паралипоменон", Shortenings = new List<string>() { "1пар", "1пр", "1 пар", "1-я паралипоменон" } },
                        new BibleBookInfo() { Name = "2-я Паралипоменон", SectionName = "14. 2-я Паралипоменон", Shortenings = new List<string>() { "2пар", "2пр", "2 пар", "2-я паралипоменон" } },
                        new BibleBookInfo() { Name = "Ездра", SectionName = "15. Ездра", Shortenings = new List<string>() { "ездр", "езд", "ез", "ездра" } },
                        new BibleBookInfo() { Name = "Неемия", SectionName = "16. Неемия", Shortenings = new List<string>() { "неем", "нм", "неемия" } },
                        new BibleBookInfo() { Name = "Есфирь", SectionName = "17. Есфирь", Shortenings = new List<string>() { "есф", "ес", "есфирь" } },
                        new BibleBookInfo() { Name = "Иов", SectionName = "18. Иов", Shortenings = new List<string>() { "иов", "ив" } },
                        new BibleBookInfo() { Name = "Псалтирь", SectionName = "19. Псалтирь", Shortenings = new List<string>() { "пс", "псалт", "псал", "псл", "псалом", "псалтырь", "псалмы", "псалтирь" } },
                        new BibleBookInfo() { Name = "Притчи", SectionName = "20. Притчи", Shortenings = new List<string>() { "прит", "притч", "при", "притчи", "притча", "пр" } },
                        new BibleBookInfo() { Name = "Екклесиаст", SectionName = "21. Екклесиаст", Shortenings = new List<string>() { "еккл", "ек", "екк", "екклесиаст" } },
                        new BibleBookInfo() { Name = "Песня Песней", SectionName = "22. Песня Песней", Shortenings = new List<string>() { "песн", "пес", "псн", "песн.песней", "песни", "песня песней" } },
                        new BibleBookInfo() { Name = "Исаия", SectionName = "23. Исаия", Shortenings = new List<string>() { "ис", "иса", "исаия" } },
                        new BibleBookInfo() { Name = "Иеремия", SectionName = "24. Иеремия", Shortenings = new List<string>() { "иер", "иерем", "иеремия" } },
                        new BibleBookInfo() { Name = "Плач Иеремии", SectionName = "25. Плач Иеремии", Shortenings = new List<string>() { "плач", "плч", "пл", "пл.иер", "плач иеремии" } },
                        new BibleBookInfo() { Name = "Иезекииль", SectionName = "26. Иезекииль", Shortenings = new List<string>() { "иез", "иезек", "иезекииль" } },
                        new BibleBookInfo() { Name = "Даниил", SectionName = "27. Даниил", Shortenings = new List<string>() { "дан", "дн", "днл", "даниил" } },
                        new BibleBookInfo() { Name = "Осия", SectionName = "28. Осия", Shortenings = new List<string>() { "ос", "осия" } },
                        new BibleBookInfo() { Name = "Иоиль", SectionName = "29. Иоиль", Shortenings = new List<string>() { "иоил", "ил", "иоиль" } },
                        new BibleBookInfo() { Name = "Амос", SectionName = "30. Амос", Shortenings = new List<string>() { "ам", "амс", "амос" } },
                        new BibleBookInfo() { Name = "Авдий", SectionName = "31. Авдий", Shortenings = new List<string>() { "авд", "авдий" } },
                        new BibleBookInfo() { Name = "Иона", SectionName = "32. Иона", Shortenings = new List<string>() { "ион", "иона" } },
                        new BibleBookInfo() { Name = "Михей", SectionName = "33. Михей", Shortenings = new List<string>() { "мих", "мх", "михей", "михея" } },
                        new BibleBookInfo() { Name = "Наум", SectionName = "34. Наум", Shortenings = new List<string>() { "наум" } },
                        new BibleBookInfo() { Name = "Аввакум", SectionName = "35. Аввакум", Shortenings = new List<string>() { "авв", "аввак", "аввакум" } },
                        new BibleBookInfo() { Name = "Софония", SectionName = "36. Софония", Shortenings = new List<string>() { "соф", "софон", "софония" } },
                        new BibleBookInfo() { Name = "Аггей", SectionName = "37. Аггей", Shortenings = new List<string>() { "агг", "аггей" } },
                        new BibleBookInfo() { Name = "Захария", SectionName = "38. Захария", Shortenings = new List<string>() { "захария", "зах", "зхр", "захар" } },
                        new BibleBookInfo() { Name = "Малахия", SectionName = "39. Малахия", Shortenings = new List<string>() { "мал", "малах", "млх", "малахия" } },
                        new BibleBookInfo() { Name = "От Матфея", SectionName = "01. От Матфея", Shortenings = new List<string>() { "мат", "матф", "мтф", "мф", "мт", "матфей", "матфея", "от матфея" } },
                        new BibleBookInfo() { Name = "От Марка", SectionName = "02. От Марка", Shortenings = new List<string>() { "мар", "марк", "мрк", "мр", "марка", "мк", "от марка" } },
                        new BibleBookInfo() { Name = "От Луки", SectionName = "03. От Луки", Shortenings = new List<string>() { "лук", "лк", "лука", "луки", "от луки" } },
                        new BibleBookInfo() { Name = "От Иоанна", SectionName = "04. От Иоанна", Shortenings = new List<string>() { "иоан", "ин", "иоанн", "иоанна", "от иоанна" } },
                        new BibleBookInfo() { Name = "Деяния", SectionName = "05. Деяния", Shortenings = new List<string>() { "деян", "дея", "д.а", "деяния", "деяния апостолов" } },
                        new BibleBookInfo() { Name = "Иакова", SectionName = "06. Иакова", Shortenings = new List<string>() { "иак", "ик", "иаков", "иакова" } },
                        new BibleBookInfo() { Name = "1-е Петра", SectionName = "07. 1-е Петра", Shortenings = new List<string>() { "1пет", "1 пет", "1пт", "1птр", "1 птр", "1петр", "1 петр", "1петра", "1-е петра", "1 петра", "1 петр", "1-е петр" } },
                        new BibleBookInfo() { Name = "2-е Петра", SectionName = "08. 2-е Петра", Shortenings = new List<string>() { "2пет", "2 пет", "2пт", "2птр", "2 птр", "2петр", "2 петр", "2петра", "2-е петра", "2 петра", "2 петр", "2-е петр" } },
                        new BibleBookInfo() { Name = "1-е Иоанна", SectionName = "09. 1-е Иоанна", Shortenings = new List<string>() { "1иоан", "1ин", "1иоанн", "1иоанна", "1-е иоанна", "1 иоанн", "1 иоанна", "1 ин", "1 иоан", "1-е иоанн", "1-е иоан", "1-е ин" } },
                        new BibleBookInfo() { Name = "2-е Иоанна", SectionName = "10. 2-е Иоанна", Shortenings = new List<string>() { "2иоан", "2ин", "2иоанн", "2иоанна", "2-е иоанна", "2 иоанн", "2 иоанна", "2 ин", "2 иоан", "2-е иоанн", "2-е иоан", "2-е ин" } },
                        new BibleBookInfo() { Name = "3-е Иоанна", SectionName = "11. 3-е Иоанна", Shortenings = new List<string>() { "3иоан", "3ин", "3иоанн", "3иоанна", "3-е иоанна", "3 иоанн", "3 иоанна", "3 ин", "3 иоан", "3-е иоанн", "3-е иоан", "3-е ин" } },
                        new BibleBookInfo() { Name = "Иуды", SectionName = "12. Иуды", Shortenings = new List<string>() { "иуд", "ид", "иуда", "иуды" } },
                        new BibleBookInfo() { Name = "К Римлянам", SectionName = "13. К Римлянам", Shortenings = new List<string>() { "рим", "римл", "римлянам", "к римлянам" } },
                        new BibleBookInfo() { Name = "1-е Коринфянам", SectionName = "14. 1-е Коринфянам", Shortenings = new List<string>() { "1кор", "1 кор", "1коринф", "1коринфянам", "1-е коринфянам", "1 коринфянам", "1 коринф", "1-е коринф", "1-е кор" } },
                        new BibleBookInfo() { Name = "2-е Коринфянам", SectionName = "15. 2-е Коринфянам", Shortenings = new List<string>() { "2кор", "2 кор", "2коринф", "2коринфянам", "2-е коринфянам", "2 коринфянам", "2 коринф", "2-е коринф", "2-е кор" } },
                        new BibleBookInfo() { Name = "К Галатам", SectionName = "16. К Галатам", Shortenings = new List<string>() { "гал", "галат", "галатам", "к галатам" } },
                        new BibleBookInfo() { Name = "К Ефесянам", SectionName = "17. К Ефесянам", Shortenings = new List<string>() { "еф", "ефес", "ефесянам", "к ефесянам" } },
                        new BibleBookInfo() { Name = "К Филиппийцам", SectionName = "18. К Филиппийцам", Shortenings = new List<string>() { "фил", "флп", "филип", "филиппийцам", "к филиппийцам" } },
                        new BibleBookInfo() { Name = "К Колоссянам", SectionName = "19. К Колоссянам", Shortenings = new List<string>() { "кол", "колос", "колоссянам", "к колоссянам" } },
                        new BibleBookInfo() { Name = "1-е Фессалоникийцам", SectionName = "20. 1-е Фессалоникийцам", Shortenings = new List<string>() { "1фесс", "1фес", "1фессалоникийцам", "1 фес", "1 фесс", "1-е фессалоникийцам", "1 фессалоникийцам", "1-е фесс", "1-е фес" } },
                        new BibleBookInfo() { Name = "2-е Фессалоникийцам", SectionName = "21. 2-е Фессалоникийцам", Shortenings = new List<string>() { "2фесс", "2фес", "2фессалоникийцам", "2 фес", "2 фесс", "2-е фессалоникийцам", "2 фессалоникийцам", "2-е фесс", "2-е фес" } },
                        new BibleBookInfo() { Name = "1-е Тимофею", SectionName = "22. 1-е Тимофею", Shortenings = new List<string>() { "1тим", "1тимоф", "1тимофею", "1 тимофею", "1-е тимофею", "1 тим", "1 тимоф", "1-е тимоф", "1-е тим" } },
                        new BibleBookInfo() { Name = "2-е Тимофею", SectionName = "23. 2-е Тимофею", Shortenings = new List<string>() { "2тим", "2тимоф", "2тимофею", "2 тимофею", "2-е тимофею", "2 тим", "2 тимоф", "2-е тимоф", "2-е тим" } },
                        new BibleBookInfo() { Name = "К Титу", SectionName = "24. К Титу", Shortenings = new List<string>() { "тит", "титу", "к титу" } },
                        new BibleBookInfo() { Name = "К Филимону", SectionName = "25. К Филимону", Shortenings = new List<string>() { "флм", "филимон", "филимону", "к филимону" } },
                        new BibleBookInfo() { Name = "К Евреям", SectionName = "26. К Евреям", Shortenings = new List<string>() { "евр", "евреям", "к евреям" } },
                        new BibleBookInfo() { Name = "Откровение", SectionName = "27. Откровение", Shortenings = new List<string>() { "откр", "отк", "откровен", "апок", "откровение", "апокалипсис" } }
                    }
                }
            };

            XmlSerializer ser = new XmlSerializer(typeof(ModuleInfo));
            using (var fs = new FileStream("c:\\manifest.xml", FileMode.Create))
            {

                ser.Serialize(fs, module);
                fs.Flush();

            }
        }
    }
}
