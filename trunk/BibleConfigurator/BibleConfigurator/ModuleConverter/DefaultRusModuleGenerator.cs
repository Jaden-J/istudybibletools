﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using System.Xml.Serialization;
using System.IO;

namespace BibleConfigurator.ModuleConverter
{
    public static class DefaultRusModuleGenerator
    {
        public static void GenerateModuleInfo(string manifestFilePath, bool addSingleNotebook)
        {
            ModuleInfo module = new ModuleInfo()
            {
                Version = "1.0",
                Name = "Синодальный перевод (Русский язык)",
                Notebooks = new List<NotebookInfo>() 
                {                    
                    new NotebookInfo() { Type = NotebookType.Bible, Name = "Библия.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleStudy, Name = "Изучение Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleComments, Name = "Комментарии к Библии.onepkg" },
                    new NotebookInfo() { Type = NotebookType.BibleNotesPages, Name = "Сводные заметок.onepkg" }
                },
                BibleStructure = new BibleStructureInfo()
                {
                    OldTestamentName = "Ветхий Завет",
                    NewTestamentName = "Новый Завет",
                    OldTestamentBooksCount = 39,
                    NewTestamentBooksCount = 27,
                    Alphabet = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯабвгдеёжзийклмнопрстуфхцчшщъыьэюя",
                    BibleBooks = new List<BibleBookInfo>()
                    {
                        new BibleBookInfo() { Name = "Бытие", SectionName = "01. Бытие", Abbreviations = new List<string>() { "быт", "бт", "бытие" } },
                        new BibleBookInfo() { Name = "Исход", SectionName = "02. Исход", Abbreviations = new List<string>() { "исх", "исход" } },
                        new BibleBookInfo() { Name = "Левит", SectionName = "03. Левит", Abbreviations = new List<string>() { "лев", "лв", "левит" } },
                        new BibleBookInfo() { Name = "Числа", SectionName = "04. Числа", Abbreviations = new List<string>() { "чис", "чс", "числ", "числа" } },
                        new BibleBookInfo() { Name = "Второзаконие", SectionName = "05. Второзаконие", Abbreviations = new List<string>() { "втор", "вт", "втрзк", "второзаконие" } },
                        new BibleBookInfo() { Name = "Иисус Навин", SectionName = "06. Иисус Навин", Abbreviations = new List<string>() { "иис.нав", "нав", "иисус навин", "ииснав", "ис. нав", "ис.нав", "навин" } },
                        new BibleBookInfo() { Name = "Судьи", SectionName = "07. Судьи", Abbreviations = new List<string>() { "суд", "сд", "судьи", "судей" } },
                        new BibleBookInfo() { Name = "Руфь", SectionName = "08. Руфь", Abbreviations = new List<string>() { "руф", "рф", "руфь" } },
                        new BibleBookInfo() { Name = "1-я Царств", SectionName = "09. 1-я Царств", Abbreviations = new List<string>() { "1цар", "1цр", "1ц", "1царств", "1-я царств" } },
                        new BibleBookInfo() { Name = "2-я Царств", SectionName = "10. 2-я Царств", Abbreviations = new List<string>() { "2цар", "2цр", "2ц", "2царств", "2-я царств" } },
                        new BibleBookInfo() { Name = "3-я Царств", SectionName = "11. 3-я Царств", Abbreviations = new List<string>() { "3цар", "3цр", "3ц", "3царств", "3-я царств" } },
                        new BibleBookInfo() { Name = "4-я Царств", SectionName = "12. 4-я Царств", Abbreviations = new List<string>() { "4цар", "4цр", "4ц", "4царств", "4-я царств" } },
                        new BibleBookInfo() { Name = "1-я Паралипоменон", SectionName = "13. 1-я Паралипоменон", Abbreviations = new List<string>() { "1пар", "1пр",  "1паралипоменон", "1-я паралипоменон" } },
                        new BibleBookInfo() { Name = "2-я Паралипоменон", SectionName = "14. 2-я Паралипоменон", Abbreviations = new List<string>() { "2пар", "2пр",  "2паралипоменон", "2-я паралипоменон" } },
                        new BibleBookInfo() { Name = "Ездра", SectionName = "15. Ездра", Abbreviations = new List<string>() { "ездр", "езд", "ез", "ездра" } },
                        new BibleBookInfo() { Name = "Неемия", SectionName = "16. Неемия", Abbreviations = new List<string>() { "неем", "нм", "неемия" } },
                        new BibleBookInfo() { Name = "Есфирь", SectionName = "17. Есфирь", Abbreviations = new List<string>() { "есф", "ес", "есфирь" } },
                        new BibleBookInfo() { Name = "Иов", SectionName = "18. Иов", Abbreviations = new List<string>() { "иов", "ив" } },
                        new BibleBookInfo() { Name = "Псалтирь", SectionName = "19. Псалтирь", Abbreviations = new List<string>() { "пс", "псалт", "псал", "псл", "псалом", "псалтырь", "псалмы", "псалтирь" } },
                        new BibleBookInfo() { Name = "Притчи", SectionName = "20. Притчи", Abbreviations = new List<string>() { "прит", "притч", "при", "притчи", "притча", "пр" } },
                        new BibleBookInfo() { Name = "Екклесиаст", SectionName = "21. Екклесиаст", Abbreviations = new List<string>() { "еккл", "ек", "екк", "екклесиаст" } },
                        new BibleBookInfo() { Name = "Песня Песней", SectionName = "22. Песня Песней", Abbreviations = new List<string>() { "песн", "пес", "псн", "песн.песней", "песни", "песня песней" } },
                        new BibleBookInfo() { Name = "Исаия", SectionName = "23. Исаия", Abbreviations = new List<string>() { "ис", "иса", "исаия", "исайя" } },
                        new BibleBookInfo() { Name = "Иеремия", SectionName = "24. Иеремия", Abbreviations = new List<string>() { "иер", "иерем", "иеремия" } },
                        new BibleBookInfo() { Name = "Плач Иеремии", SectionName = "25. Плач Иеремии", Abbreviations = new List<string>() { "плач", "плч", "пл", "пл.иер", "плач иеремии" } },
                        new BibleBookInfo() { Name = "Иезекииль", SectionName = "26. Иезекииль", Abbreviations = new List<string>() { "иез", "иезек", "иезекииль" } },
                        new BibleBookInfo() { Name = "Даниил", SectionName = "27. Даниил", Abbreviations = new List<string>() { "дан", "дн", "днл", "даниил" } },
                        new BibleBookInfo() { Name = "Осия", SectionName = "28. Осия", Abbreviations = new List<string>() { "ос", "осия" } },
                        new BibleBookInfo() { Name = "Иоиль", SectionName = "29. Иоиль", Abbreviations = new List<string>() { "иоил", "ил", "иоиль" } },
                        new BibleBookInfo() { Name = "Амос", SectionName = "30. Амос", Abbreviations = new List<string>() { "ам", "амс", "амос" } },
                        new BibleBookInfo() { Name = "Авдий", SectionName = "31. Авдий", Abbreviations = new List<string>() { "авд", "авдий" } },
                        new BibleBookInfo() { Name = "Иона", SectionName = "32. Иона", Abbreviations = new List<string>() { "ион", "иона" } },
                        new BibleBookInfo() { Name = "Михей", SectionName = "33. Михей", Abbreviations = new List<string>() { "мих", "мх", "михей", "михея" } },
                        new BibleBookInfo() { Name = "Наум", SectionName = "34. Наум", Abbreviations = new List<string>() { "наум" } },
                        new BibleBookInfo() { Name = "Аввакум", SectionName = "35. Аввакум", Abbreviations = new List<string>() { "авв", "аввак", "аввакум" } },
                        new BibleBookInfo() { Name = "Софония", SectionName = "36. Софония", Abbreviations = new List<string>() { "соф", "софон", "софония" } },
                        new BibleBookInfo() { Name = "Аггей", SectionName = "37. Аггей", Abbreviations = new List<string>() { "агг", "аггей" } },
                        new BibleBookInfo() { Name = "Захария", SectionName = "38. Захария", Abbreviations = new List<string>() { "захария", "зах", "зхр", "захар" } },
                        new BibleBookInfo() { Name = "Малахия", SectionName = "39. Малахия", Abbreviations = new List<string>() { "мал", "малах", "млх", "малахия" } },
                        new BibleBookInfo() { Name = "От Матфея", SectionName = "01. От Матфея", Abbreviations = new List<string>() { "мат", "матф", "мтф", "мф", "мт", "матфей", "матфея", "от матфея" } },
                        new BibleBookInfo() { Name = "От Марка", SectionName = "02. От Марка", Abbreviations = new List<string>() { "мар", "марк", "мрк", "мр", "марка", "мк", "от марка" } },
                        new BibleBookInfo() { Name = "От Луки", SectionName = "03. От Луки", Abbreviations = new List<string>() { "лук", "лк", "лука", "луки", "от луки" } },
                        new BibleBookInfo() { Name = "От Иоанна", SectionName = "04. От Иоанна", Abbreviations = new List<string>() { "иоан", "ин", "иоанн", "иоанна", "от иоанна" } },
                        new BibleBookInfo() { Name = "Деяния", SectionName = "05. Деяния", Abbreviations = new List<string>() { "деян", "дея", "д.а", "деяния", "деяния апостолов" } },
                        new BibleBookInfo() { Name = "Иакова", SectionName = "06. Иакова", Abbreviations = new List<string>() { "иак", "ик", "иаков", "иакова" } },
                        new BibleBookInfo() { Name = "1-е Петра", SectionName = "07. 1-е Петра", Abbreviations = new List<string>() { "1пет", "1пт", "1птр", "1петр", "1петра", "1-е петра", "1-е петр" } },
                        new BibleBookInfo() { Name = "2-е Петра", SectionName = "08. 2-е Петра", Abbreviations = new List<string>() { "2пет",  "2пт", "2птр","2петр", "2петра", "2-е петра", "2-е петр" } },
                        new BibleBookInfo() { Name = "1-е Иоанна", SectionName = "09. 1-е Иоанна", Abbreviations = new List<string>() { "1иоан", "1ин", "1иоанн", "1иоанна", "1-е иоанна", "1-е иоанн", "1-е иоан", "1-е ин" } },
                        new BibleBookInfo() { Name = "2-е Иоанна", SectionName = "10. 2-е Иоанна", Abbreviations = new List<string>() { "2иоан", "2ин", "2иоанн", "2иоанна", "2-е иоанна", "2-е иоанн", "2-е иоан", "2-е ин" } },
                        new BibleBookInfo() { Name = "3-е Иоанна", SectionName = "11. 3-е Иоанна", Abbreviations = new List<string>() { "3иоан", "3ин", "3иоанн", "3иоанна", "3-е иоанна", "3-е иоанн", "3-е иоан", "3-е ин" } },
                        new BibleBookInfo() { Name = "Иуды", SectionName = "12. Иуды", Abbreviations = new List<string>() { "иуд", "ид", "иуда", "иуды" } },
                        new BibleBookInfo() { Name = "К Римлянам", SectionName = "13. К Римлянам", Abbreviations = new List<string>() { "рим", "римл", "римлянам", "к римлянам" } },
                        new BibleBookInfo() { Name = "1-е Коринфянам", SectionName = "14. 1-е Коринфянам", Abbreviations = new List<string>() { "1кор", "1коринф", "1коринфянам", "1-е коринфянам", "1-е коринф", "1-е кор" } },
                        new BibleBookInfo() { Name = "2-е Коринфянам", SectionName = "15. 2-е Коринфянам", Abbreviations = new List<string>() { "2кор", "2коринф", "2коринфянам", "2-е коринфянам", "2-е коринф", "2-е кор" } },
                        new BibleBookInfo() { Name = "К Галатам", SectionName = "16. К Галатам", Abbreviations = new List<string>() { "гал", "галат", "галатам", "к галатам" } },
                        new BibleBookInfo() { Name = "К Ефесянам", SectionName = "17. К Ефесянам", Abbreviations = new List<string>() { "еф", "ефес", "ефесянам", "к ефесянам" } },
                        new BibleBookInfo() { Name = "К Филиппийцам", SectionName = "18. К Филиппийцам", Abbreviations = new List<string>() { "фил", "флп", "филип", "филиппийцам", "к филиппийцам" } },
                        new BibleBookInfo() { Name = "К Колоссянам", SectionName = "19. К Колоссянам", Abbreviations = new List<string>() { "кол", "колос", "колоссянам", "к колоссянам" } },
                        new BibleBookInfo() { Name = "1-е Фессалоникийцам", SectionName = "20. 1-е Фессалоникийцам", Abbreviations = new List<string>() { "1фесс", "1фес", "1фессалоникийцам", "1-е фессалоникийцам", "1-е фесс", "1-е фес" } },
                        new BibleBookInfo() { Name = "2-е Фессалоникийцам", SectionName = "21. 2-е Фессалоникийцам", Abbreviations = new List<string>() { "2фесс", "2фес", "2фессалоникийцам", "2-е фессалоникийцам", "2-е фесс", "2-е фес" } },
                        new BibleBookInfo() { Name = "1-е Тимофею", SectionName = "22. 1-е Тимофею", Abbreviations = new List<string>() { "1тим", "1тимоф", "1тимофею", "1-е тимофею", "1-е тимоф", "1-е тим" } },
                        new BibleBookInfo() { Name = "2-е Тимофею", SectionName = "23. 2-е Тимофею", Abbreviations = new List<string>() { "2тим", "2тимоф", "2тимофею", "2-е тимофею", "2-е тимоф", "2-е тим" } },
                        new BibleBookInfo() { Name = "К Титу", SectionName = "24. К Титу", Abbreviations = new List<string>() { "тит", "титу", "к титу" } },
                        new BibleBookInfo() { Name = "К Филимону", SectionName = "25. К Филимону", Abbreviations = new List<string>() { "флм", "филимон", "филимону", "к филимону" } },
                        new BibleBookInfo() { Name = "К Евреям", SectionName = "26. К Евреям", Abbreviations = new List<string>() { "евр", "евреям", "к евреям" } },
                        new BibleBookInfo() { Name = "Откровение", SectionName = "27. Откровение", Abbreviations = new List<string>() { "откр", "отк", "откровен", "апок", "откровение", "апокалипсис" } }
                    }
                }
            };

            if (addSingleNotebook)
                module.Notebooks.Add(new NotebookInfo()
                {
                    Type = NotebookType.Single,
                    Name = "Holy Bible.onepkg",
                    SectionGroups = new List<BibleCommon.Common.SectionGroupInfo>()
                    {
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.Bible, Name="Библия" },
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleStudy, Name="Изучение Библии" },
                        new BibleCommon.Common.SectionGroupInfo() { Type = SectionGroupType.BibleComments, Name="Комментарии к Библии" }                        
                    }
                });

            XmlSerializer ser = new XmlSerializer(typeof(ModuleInfo));
            using (var fs = new FileStream(manifestFilePath, FileMode.Create))
            {
                ser.Serialize(fs, module);
                fs.Flush();
            }
        }
    }
}