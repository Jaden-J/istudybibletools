using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections.Specialized;

namespace BibleCommon.Consts
{
    public static class Constants
    {
        public static readonly string OneNoteXmlNs = "http://schemas.microsoft.com/office/onenote/2010/onenote";
        public static readonly string ToolsName = "IStudyBibleTools";
        public static readonly string ConfigFileName = "settings.config";

        public static readonly string DefaultNotebookNameBible = "Библия";
        public static readonly string DefaultNotebookNameBibleComments = "Комментарии к Библии";
        public static readonly string DefaultNotebookNameBibleStudy = "Изучение Библии";
        public static readonly string DefaultPageNameDefaultBookOverview = "Общий обзор";
        public static readonly string DefaultPageNameDefaultComments = "Комментарии";
        public static readonly string DefaultPageName_Notes = "Заметки";

        public static readonly string ParameterName_NotebookIdBible = "NotebookId_Bible";
        public static readonly string ParameterName_NotebookIdBibleComments = "NotebookId_BibleComments";
        public static readonly string ParameterName_NotebookIdBibleStudy = "NotebookId_BibleStudy";        
        public static readonly string ParameterName_SectionGroupIdBible = "SectionGroupId_Bible";
        public static readonly string ParameterName_SectionGroupIdBibleComments = "SectionGroupId_BibleComments";
        public static readonly string ParameterName_SectionGroupIdBibleStudy = "SectionGroupId_BibleStudy";
        public static readonly string ParameterName_PageNameDefaultComments = "PageName_DefaultComments";
        public static readonly string ParameterName_PageNameDefaultBookOverview = "PageName_DefaultBookOverview";
        public static readonly string ParameterName_PageNamePageName_Notes = "PageName_Notes";
        public static readonly string ParameterName_LastNotesLinkTime = "LastNotesLinkTime";
        public static readonly string ParameterName_NewVersionOnServer = "NewVersionOnServer";
        public static readonly string ParameterName_NewVersionOnServerLatestCheckTime = "NewVersionOnServerLatestCheckTime";

        public static readonly string NewVersionOnServerFileUrl = "http://IStudyBibleTools.ru/ServerVariables.xml";
        public static readonly TimeSpan NewVersionCheckPeriod = new TimeSpan(1, 0, 0, 0);

        public static readonly string DownloadPageUrl = "http://IStudyBibleTools.ru/download.htm";

        public static readonly string Error_SystemIsNotConfigures = "Программа не сконфигурирована";

        public static readonly string VerseLinkTemplate = "ссылка {0}";
        public static readonly string VerseLinksDelimiter = "; ";
        


        public static readonly Dictionary<string, StringCollection> BookNames = new Dictionary<string, StringCollection>()
            {
                { "01. Бытие", new StringCollection() { "быт", "бт", "бытие" } },
                { "02. Исход", new StringCollection() { "исх", "исход" } },
                { "03. Левит", new StringCollection() { "лев", "лв", "левит" } },
                { "04. Числа", new StringCollection() { "чис", "чс", "числ", "числа" } },
                { "05. Второзаконие", new StringCollection() { "втор", "вт", "втрзк", "второзаконие" } },
                { "06. Иисус Навин", new StringCollection() { "иис.нав", "нав", "иисус навин", "ииснав", "ис. нав", "ис.нав", "навин" } },
                { "07. Судьи", new StringCollection() { "суд", "сд", "судьи", "судей" } },
                { "08. Руфь", new StringCollection() { "руф", "рф", "руфь" } },
                { "09. 1-я Царств", new StringCollection() { "1цар", "1 цар", "1цр", "1ц", "1царств", "1-я царств", "1 царств" } },
                { "10. 2-я Царств", new StringCollection() { "2цар", "2 цар", "2цр", "2ц", "2царств", "2-я царств", "2 царств" } },
                { "11. 3-я Царств", new StringCollection() { "3цар", "3 цар", "3цр", "3ц", "3царств", "3-я царств", "3 царств" } },
                { "12. 4-я Царств", new StringCollection() { "4цар", "4 цар", "4цр", "4ц", "4царств", "4-я царств", "4 царств" } },
                { "13. 1-я Паралипоменон", new StringCollection() { "1пар", "1пр", "1 пар", "1-я паралипоменон" } },
                { "14. 2-я Паралипоменон", new StringCollection() { "2пар", "2пр", "2 пар", "2-я паралипоменон" } },
                { "15. Ездра", new StringCollection() { "ездр", "езд", "ез", "ездра" } },
                { "16. Неемия", new StringCollection() { "неем", "нм", "неемия" } },
                { "17. Есфирь", new StringCollection() { "есф", "ес", "есфирь" } },
                { "18. Иов", new StringCollection() { "иов", "ив" } },
                { "19. Псалтирь", new StringCollection() { "пс", "псалт", "псал", "псл", "псалом", "псалтырь", "псалмы", "псалтирь" } },
                { "20. Притчи", new StringCollection() { "прит", "притч", "при", "притчи", "притча", "пр" } },
                { "21. Екклесиаст", new StringCollection() { "еккл", "ек", "екк", "екклесиаст" } },
                { "22. Песня Песней", new StringCollection() { "песн", "пес", "псн", "песн.песней", "песни", "песня песней" } },
                { "23. Исаия", new StringCollection() { "ис", "иса", "исаия" } },
                { "24. Иеремия", new StringCollection() { "иер", "иерем", "иеремия" } },
                { "25. Плач Иеремии", new StringCollection() { "плач", "плч", "пл", "пл.иер", "плач иеремии" } },
                { "26. Иезекииль", new StringCollection() { "иез", "иезек", "иезекииль" } },
                { "27. Даниил", new StringCollection() { "дан", "дн", "днл", "даниил" } },
                { "28. Осия", new StringCollection() { "ос", "осия" } },
                { "29. Иоиль", new StringCollection() { "иоил", "ил", "иоиль" } },
                { "30. Амос", new StringCollection() { "ам", "амс", "амос" } },
                { "31. Авдий", new StringCollection() { "авд", "авдий" } },
                { "32. Иона", new StringCollection() { "ион", "иона" } },
                { "33. Михей", new StringCollection() { "мих", "мх", "михей", "михея" } },
                { "34. Наум", new StringCollection() { "наум" } },
                { "35. Аввакум", new StringCollection() { "авв", "аввак", "аввакум" } },
                { "36. Софония", new StringCollection() { "соф", "софон", "софония" } },
                { "37. Аггей", new StringCollection() { "агг", "аггей" } },
                { "38. Захария", new StringCollection() { "захария", "зах", "зхр", "захар" } },
                { "39. Малахия", new StringCollection() { "мал", "малах", "млх", "малахия" } },
                { "01. От Матфея", new StringCollection() { "мат", "матф", "мтф", "мф", "мт", "матфей", "матфея", "от матфея" } },
                { "02. От Марка", new StringCollection() { "мар", "марк", "мрк", "мр", "марка", "мк", "от марка" } },
                { "03. От Луки", new StringCollection() { "лук", "лк", "лука", "луки", "от луки" } },
                { "04. От Иоанна", new StringCollection() { "иоан", "ин", "иоанн", "иоанна", "от иоанна" } },
                { "05. Деяния", new StringCollection() { "деян", "дея", "д.а", "деяния", "деяния апостолов" } },
                { "06. Иакова", new StringCollection() { "иак", "ик", "иаков", "иакова" } },
                { "07. 1-е Петра", new StringCollection() { "1пет", "1 пет", "1пт", "1птр", "1 птр", "1петр", "1 петр", "1петра", "1-е петра", "1 петра", "1 петр", "1-е петр" } },
                { "08. 2-е Петра", new StringCollection() { "2пет", "2 пет", "2пт", "2птр", "2 птр", "2петр", "2 петр", "2петра", "2-е петра", "2 петра", "2 петр", "2-е петр" } },
                { "09. 1-е Иоанна", new StringCollection() { "1иоан", "1ин", "1иоанн", "1иоанна", "1-е иоанна", "1 иоанн", "1 иоанна", "1 ин", "1 иоан", "1-е иоанн", "1-е иоан", "1-е ин" } },
                { "10. 2-е Иоанна", new StringCollection() { "2иоан", "2ин", "2иоанн", "2иоанна", "2-е иоанна", "2 иоанн", "2 иоанна", "2 ин", "2 иоан", "2-е иоанн", "2-е иоан", "2-е ин" } },
                { "11. 3-е Иоанна", new StringCollection() { "3иоан", "3ин", "3иоанн", "3иоанна", "3-е иоанна", "3 иоанн", "3 иоанна", "3 ин", "3 иоан", "3-е иоанн", "3-е иоан", "3-е ин" } },
                { "12. Иуды", new StringCollection() { "иуд", "ид", "иуда", "иуды" } },
                { "13. К Римлянам", new StringCollection() { "рим", "римл", "римлянам", "к римлянам" } },
                { "14. 1-е Коринфянам", new StringCollection() { "1кор", "1 кор", "1коринф", "1коринфянам", "1-е коринфянам", "1 коринфянам", "1 коринф", "1-е коринф", "1-е кор" } },
                { "15. 2-е Коринфянам", new StringCollection() { "2кор", "2 кор", "2коринф", "2коринфянам", "2-е коринфянам", "2 коринфянам", "2 коринф", "2-е коринф", "2-е кор" } },
                { "16. К Галатам", new StringCollection() { "гал", "галат", "галатам", "к галатам" } },
                { "17. К Ефесянам", new StringCollection() { "еф", "ефес", "ефесянам", "к ефесянам" } },
                { "18. К Филиппийцам", new StringCollection() { "фил", "флп", "филип", "филиппийцам", "к филиппийцам" } },
                { "19. К Колоссянам", new StringCollection() { "кол", "колос", "колоссянам", "к колоссянам" } },
                { "20. 1-е Фессалоникийцам", new StringCollection() { "1фесс", "1фес", "1фессалоникийцам", "1 фес", "1 фесс", "1-е фессалоникийцам", "1 фессалоникийцам", "1-е фесс", "1-е фес" } },
                { "21. 2-е Фессалоникийцам", new StringCollection() { "2фесс", "2фес", "2фессалоникийцам", "2 фес", "2 фесс", "2-е фессалоникийцам", "2 фессалоникийцам", "2-е фесс", "2-е фес" } },
                { "22. 1-е Тимофею", new StringCollection() { "1тим", "1тимоф", "1тимофею", "1 тимофею", "1-е тимофею", "1 тим", "1 тимоф", "1-е тимоф", "1-е тим" } },
                { "23. 2-е Тимофею", new StringCollection() { "2тим", "2тимоф", "2тимофею", "2 тимофею", "2-е тимофею", "2 тим", "2 тимоф", "2-е тимоф", "2-е тим" } },
                { "24. К Титу", new StringCollection() { "тит", "титу", "к титу" } },
                { "25. К Филимону", new StringCollection() { "флм", "филимон", "филимону", "к филимону" } },
                { "26. К Евреям", new StringCollection() { "евр", "евреям", "к евреям" } },
                { "27. Откровение", new StringCollection() { "откр", "отк", "откровен", "апок", "откровение", "апокалипсис" } }
            };


    }
}
