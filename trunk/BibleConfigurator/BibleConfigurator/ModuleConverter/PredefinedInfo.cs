using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using BibleCommon.Common;
using BibleCommon.Helpers;

namespace BibleConfigurator.ModuleConverter
{
    public static class PredefinedNotebooksInfo
    {   
        public static List<NotebookInfo> RussianStrong
        {
            get
            {
                return new List<NotebookInfo>() 
                {  
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.Bible, 
                        Name = "Библия.onepkg", 
                        SkipCheck = true,
                        SectionGroups = RussianNotebookBibleSectionGroups
                    }
                };
            }
        }

        public static List<NotebookInfo> Russian
        {
            get
            {
                return new List<NotebookInfo>() 
                {  
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.Bible, 
                        Name = "Библия.onepkg", 
                        SectionGroups = RussianNotebookBibleSectionGroups
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleStudy, 
                        Name = "Изучение Библии.onepkg" 
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleComments, 
                        Name = "Комментарии к Библии.onepkg",
                        SectionGroups = RussianNotebookCommentsSectionGroups
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleNotesPages, 
                        Name = "Сводные заметок.onepkg",
                        SectionGroups = RussianNotebookCommentsSectionGroups
                    }
                };
            }
        }

        public static List<NotebookInfo> Russian77
        {
            get
            {
                return new List<NotebookInfo>() 
                {  
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.Bible, 
                        Name = "Библия.onepkg", 
                        SectionGroups = RussianNotebookBibleSectionGroups77
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleStudy, 
                        Name = "Изучение Библии.onepkg" 
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleComments, 
                        Name = "Комментарии к Библии.onepkg",
                        SectionGroups = RussianNotebookCommentsSectionGroups
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleNotesPages, 
                        Name = "Сводные заметок.onepkg",
                        SectionGroups = RussianNotebookCommentsSectionGroups
                    }
                };
            }
        }

        private static List<SectionGroupInfo> RussianNotebookBibleSectionGroups
        {
            get
            {
                return new List<SectionGroupInfo>() 
                        {        
                            new SectionGroupInfo() 
                            { 
                                Name = "Ветхий Завет", 
                                CheckSectionsCount = true, 
                                SectionsCount = 39, 
                                Type = ContainerType.OldTestament
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "Новый Завет", 
                                CheckSectionsCount = true, 
                                SectionsCount = 27, 
                                Type = ContainerType.NewTestament 
                            }
                        };
            }
        }

        private static List<SectionGroupInfo> RussianNotebookBibleSectionGroups77
        {
            get
            {
                return new List<SectionGroupInfo>() 
                        {        
                            new SectionGroupInfo() 
                            { 
                                Name = "Ветхий Завет", 
                                CheckSectionsCount = true, 
                                SectionsCount = 50, 
                                Type = ContainerType.OldTestament
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "Новый Завет", 
                                CheckSectionsCount = true, 
                                SectionsCount = 27, 
                                Type = ContainerType.NewTestament 
                            }
                        };
            }
        }

        private static List<SectionGroupInfo> RussianNotebookCommentsSectionGroups
        {
            get
            {
                return new List<SectionGroupInfo>()
                        {
                            new SectionGroupInfo() 
                            { 
                                Name = "Ветхий Завет", 
                                CheckSectionsCount = true, 
                                SectionsCountMax = 3, 
                                Type = ContainerType.OldTestament 
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "Новый Завет", 
                                CheckSectionsCount = true, 
                                SectionsCountMax = 3, 
                                Type = ContainerType.NewTestament 
                            }
                        };
            }
        }



        public static List<NotebookInfo> EnglishStrong
        {
            get
            {
                return new List<NotebookInfo>()             
                {   
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.Bible, 
                        Name = "Bible.onepkg",
                        SkipCheck = true,
                        SectionGroups = EnglishNotebookBibleSectionGroups
                    }
                };
            }
        }


        public static List<NotebookInfo> English
        {
            get
            {
                return new List<NotebookInfo>()             
                {   
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.Bible, 
                        Name = "Bible.onepkg",
                        SectionGroups = EnglishNotebookBibleSectionGroups
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleStudy, 
                        Name = "Bible Study.onepkg" 
                    },
                    new NotebookInfo() 
                    { 
                        Type = ContainerType.BibleComments, 
                        Name = "Comments to the Bible.onepkg",
                        SectionGroups = EnglishNotebookCommentsSectionGroups
                    },
                    new NotebookInfo()
                    { 
                        Type = ContainerType.BibleNotesPages, 
                        Name = "Summary of Notes.onepkg",
                        SectionGroups = EnglishNotebookCommentsSectionGroups 
                    }
                };
            }
        }

        private static List<SectionGroupInfo> EnglishNotebookBibleSectionGroups
        {
            get
            {
                return new List<SectionGroupInfo>() 
                        {        
                            new SectionGroupInfo() 
                            { 
                                Name = "1. Old Testament", 
                                CheckSectionsCount = true, 
                                SectionsCount = 39, 
                                Type = ContainerType.OldTestament
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "2. New Testament", 
                                CheckSectionsCount = true, 
                                SectionsCount = 27, 
                                Type = ContainerType.NewTestament 
                            }
                        };
            }
        }

        private static List<SectionGroupInfo> EnglishNotebookCommentsSectionGroups
        {
            get
            {
                return new List<SectionGroupInfo>()
                        {
                            new SectionGroupInfo() 
                            { 
                                Name = "1. Old Testament", 
                                CheckSectionsCount = true, 
                                SectionsCountMax = 3, 
                                Type = ContainerType.OldTestament 
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "2. New Testament", 
                                CheckSectionsCount = true, 
                                SectionsCountMax = 3, 
                                Type = ContainerType.NewTestament 
                            }
                        };
            }
        }
    }

    public static class PredefinedBookIndexes
    {
        private static IEnumerable<int> GetRange(int min, int max)
        {
            for (int i = min; i <= max; i++)
                yield return i;
        }
        
        public static List<int> KJV
        {
            get
            {
                return new List<int>(GetRange(1, 66));                
            }
        }

        
        public static List<int> RST
        {
            get
            {
                var _rst = new List<int>(GetRange(1, 44));
                _rst.AddRange(GetRange(59, 65));
                _rst.AddRange(GetRange(45, 58));
                _rst.Add(66);

                return _rst;
            }
        }

        public static List<int> RST77
        {
            get
            {
                var _rst = new List<int>(GetRange(1, 39));
                _rst.AddRange(GetRange(67, 77));
                _rst.AddRange(GetRange(40, 44));
                _rst.AddRange(GetRange(59, 65));
                _rst.AddRange(GetRange(45, 58));
                _rst.Add(66);

                return _rst;
            }
        }
    }

    public static class PredefinedSectionsInfo
    {
        public static List<SectionInfo> None
        {
            get
            {
                return null;
            }
        }

        public static List<SectionInfo> RSTStrong
        {
            get
            {
                return new List<SectionInfo>()
                {
                    new SectionInfo() { Name = "Ветхий Завет.one" },
                    new SectionInfo() { Name = "Новый Завет.one" }
                };
            }
        }
    }

    //public static class PredefinedBookDifferences
    //{
    //    //public static readonly BibleTranslationDifferences KJV = new BibleTranslationDifferences();        

    //    //private static BibleTranslationDifferences _rst;
    //    //public static BibleTranslationDifferences RST
    //    //{
    //    //    get
    //    //    {
    //    //        if (_rst == null)
    //    //        {
    //    //            _rst = new BibleTranslationDifferences();
    //    //            _rst.PartVersesAlphabet = "абвгд";
    //    //            _rst.BookDifferences.AddRange(new List<BibleBookDifferences>()
    //    //            {
    //    //                new BibleBookDifferences(3, 
    //    //                            new BibleBookDifference("14:55-56", "14:55"),
    //    //                            new BibleBookDifference("14:57", "14:56")),
    //    //                new BibleBookDifferences(4,
    //    //                            new BibleBookDifference("12:16", "13:1"),
    //    //                            new BibleBookDifference("13:1-33", "13:X+1"),
    //    //                            new BibleBookDifference("29:40", "30:1"),
    //    //                            new BibleBookDifference("30:1-16", "30:X+1")),
    //    //                new BibleBookDifferences(6,
    //    //                            new BibleBookDifference("6:1", "5:16"),
    //    //                            new BibleBookDifference("6:2-27", "6:X-1")),
    //    //                new BibleBookDifferences(9,
    //    //                            new BibleBookDifference("20:42", "20:42-43"),
    //    //                            new BibleBookDifference("23:29", "24:1"),
    //    //                            new BibleBookDifference("24:1-22", "24:X+1")),
    //    //                new BibleBookDifferences(18,
    //    //                            new BibleBookDifference("40:1-5", "39:31-35"),
    //    //                            new BibleBookDifference("40:6-24", "40:X-5"),
    //    //                            new BibleBookDifference("41:1-8", "40:20-27"),
    //    //                            new BibleBookDifference("41:9-34", "41:X-8")),
    //    //                new BibleBookDifferences(19,
    //    //                            new BibleBookDifference("3:1", "3:1-2"),
    //    //                            new BibleBookDifference("3:2-8", "X:X+1"),
    //    //                            new BibleBookDifference("4:1", "4:1-2"),
    //    //                            new BibleBookDifference("4:2-8", "X:X+1"),
    //    //                            new BibleBookDifference("5:1", "5:1-2"),
    //    //                            new BibleBookDifference("5:2-12", "X:X+1"),
    //    //                            new BibleBookDifference("6:1", "6:1-2"),
    //    //                            new BibleBookDifference("6:2-10", "X:X+1"),
    //    //                            new BibleBookDifference("7:1", "7:1-2"),
    //    //                            new BibleBookDifference("7:2-17", "X:X+1"),
    //    //                            new BibleBookDifference("8:1", "8:1-2"),
    //    //                            new BibleBookDifference("8:2-9", "X:X+1"),
    //    //                            new BibleBookDifference("9:1", "9:1-2"),
    //    //                            new BibleBookDifference("9:2-20", "X:X+1"),
    //    //                            new BibleBookDifference("10:1", "9:22"),
    //    //                            new BibleBookDifference("10:2-18", "9:X+21"),
    //    //                            new BibleBookDifference("11:1-7", "10:X"),
    //    //                            new BibleBookDifference("12:1", "X-1:1-2"),
    //    //                            new BibleBookDifference("12:2-8", "X-1:X+1"),
    //    //                            new BibleBookDifference("13:1-4", "X-1:X+1"),
    //    //                            new BibleBookDifference("13:5-6", "12:6"),
    //    //                            new BibleBookDifference("14:1-7", "X-1:X"),
    //    //                            new BibleBookDifference("15:1-5", "X-1:X"),
    //    //                            new BibleBookDifference("16:1-11", "X-1:X"),
    //    //                            new BibleBookDifference("17:1-15", "X-1:X"),
    //    //                            new BibleBookDifference("18:1-50", "X-1:X+1"),
    //    //                            new BibleBookDifference("19:1-14", "X-1:X+1"),
    //    //                            new BibleBookDifference("20:1-9", "X-1:X+1"),
    //    //                            new BibleBookDifference("21:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("22:1-31", "X-1:X+1"),
    //    //                            new BibleBookDifference("23:1-6", "X-1:X"),
    //    //                            new BibleBookDifference("24:1-10", "X-1:X"),
    //    //                            new BibleBookDifference("25:1-22", "X-1:X"),
    //    //                            new BibleBookDifference("26:1-12", "X-1:X"),
    //    //                            new BibleBookDifference("27:1-14", "X-1:X"),
    //    //                            new BibleBookDifference("28:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("29:1-11", "X-1:X"),
    //    //                            new BibleBookDifference("30:1-12", "X-1:X+1"),
    //    //                            new BibleBookDifference("31:1-24", "X-1:X+1"),
    //    //                            new BibleBookDifference("32:1-11", "X-1:X"),
    //    //                            new BibleBookDifference("33:1-22", "X-1:X"),
    //    //                            new BibleBookDifference("34:1-22", "33:X+1"),
    //    //                            new BibleBookDifference("35:1-28", "34:X"),
    //    //                            new BibleBookDifference("36:1-12", "35:X+1"),
    //    //                            new BibleBookDifference("37:1-40", "36:X"),
    //    //                            new BibleBookDifference("38:1-22", "X-1:X+1"),
    //    //                            new BibleBookDifference("39:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("40:1-17", "X-1:X+1"),
    //    //                            new BibleBookDifference("41:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("42:1-11", "X-1:X+1"),
    //    //                            new BibleBookDifference("43:1-5", "42:X"),
    //    //                            new BibleBookDifference("44:1-26", "X-1:X+1"),
    //    //                            new BibleBookDifference("45:1-17", "X-1:X+1"),
    //    //                            new BibleBookDifference("46:1-11", "X-1:X+1"),
    //    //                            new BibleBookDifference("47:1-9", "X-1:X+1"),
    //    //                            new BibleBookDifference("48:1-14", "X-1:X+1"),
    //    //                            new BibleBookDifference("49:1-20", "X-1:X+1"),
    //    //                            new BibleBookDifference("50:1-23", "49:X"),
    //    //                            new BibleBookDifference("51:1-19", "X-1:X+2"),
    //    //                            new BibleBookDifference("52:1-9", "X-1:X+2"),
    //    //                            new BibleBookDifference("53:1-6", "52:X+1"),
    //    //                            new BibleBookDifference("54:1-7", "53:X+2"),
    //    //                            new BibleBookDifference("55:1-23", "X-1:X+1"),
    //    //                            new BibleBookDifference("56:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("57:1-11", "X-1:X+1"),
    //    //                            new BibleBookDifference("58:1-11", "X-1:X+1"),
    //    //                            new BibleBookDifference("59:1-17", "X-1:X+1"),
    //    //                            new BibleBookDifference("60:1-12", "59:X+2"),
    //    //                            new BibleBookDifference("61:1-8", "X-1:X+1"),
    //    //                            new BibleBookDifference("62:1-12", "X-1:X+1"),
    //    //                            new BibleBookDifference("63:1-11", "X-1:X+1"),
    //    //                            new BibleBookDifference("64:1-10", "X-1:X+1"),
    //    //                            new BibleBookDifference("65:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("66:1-20", "65:X"),
    //    //                            new BibleBookDifference("67:1-7", "X-1:X+1"),
    //    //                            new BibleBookDifference("68:1-35", "X-1:X+1"),
    //    //                            new BibleBookDifference("69:1-36", "X-1:X+1"),
    //    //                            new BibleBookDifference("70:1-5", "X-1:X+1"),
    //    //                            new BibleBookDifference("71:1-24", "X-1:X"),
    //    //                            new BibleBookDifference("72:1-20", "X-1:X"),
    //    //                            new BibleBookDifference("73:1-28", "X-1:X"),
    //    //                            new BibleBookDifference("74:1-23", "X-1:X"),
    //    //                            new BibleBookDifference("75:1-10", "X-1:X+1"),
    //    //                            new BibleBookDifference("76:1-12", "X-1:X+1"),
    //    //                            new BibleBookDifference("77:1-20", "X-1:X+1"),
    //    //                            new BibleBookDifference("78:1-72", "X-1:X"),
    //    //                            new BibleBookDifference("79:1-13", "X-1:X"),
    //    //                            new BibleBookDifference("80:1-19", "X-1:X+1"),
    //    //                            new BibleBookDifference("81:1-16", "X-1:X+1"),
    //    //                            new BibleBookDifference("82:1-8", "81:X"),
    //    //                            new BibleBookDifference("83:1-18", "X-1:X+1"),
    //    //                            new BibleBookDifference("84:1-12", "X-1:X+1"),
    //    //                            new BibleBookDifference("85:1-13", "X-1:X+1"),
    //    //                            new BibleBookDifference("86:1-17", "85:X"),
    //    //                            new BibleBookDifference("87:1-2", "86:2"),
    //    //                            new BibleBookDifference("87:3-7", "86:X"),
    //    //                            new BibleBookDifference("88:1-18", "X-1:X+1"),
    //    //                            new BibleBookDifference("89:1-52", "X-1:X+1"),
    //    //                            new BibleBookDifference("90:1-4", "X-1:X+1"),
    //    //                            new BibleBookDifference("90:5-6", "89:6"),
    //    //                            new BibleBookDifference("90:7-17", "X-1:X"),
    //    //                            new BibleBookDifference("91:1-16", "X-1:X"),
    //    //                            new BibleBookDifference("92:1-15", "91:X+1"),
    //    //                            new BibleBookDifference("93:1-5", "X-1:X"),
    //    //                            new BibleBookDifference("94:1-23", "X-1:X"),
    //    //                            new BibleBookDifference("95:1-11", "X-1:X"),
    //    //                            new BibleBookDifference("96:1-13", "X-1:X"),
    //    //                            new BibleBookDifference("97:1-12", "X-1:X"),
    //    //                            new BibleBookDifference("98:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("99:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("100:1-5", "X-1:X"),
    //    //                            new BibleBookDifference("101:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("102:1-28", "101:X+1"),
    //    //                            new BibleBookDifference("103:1-22", "X-1:X"),
    //    //                            new BibleBookDifference("104:1-35", "X-1:X"),
    //    //                            new BibleBookDifference("105:1-45", "X-1:X"),
    //    //                            new BibleBookDifference("106:1-48", "X-1:X"),
    //    //                            new BibleBookDifference("107:1-43", "X-1:X"),
    //    //                            new BibleBookDifference("108:1-13", "107:X+1"),
    //    //                            new BibleBookDifference("109:1-31", "X-1:X"),
    //    //                            new BibleBookDifference("110:1-7", "X-1:X"),
    //    //                            new BibleBookDifference("111:1-10", "X-1:X"),
    //    //                            new BibleBookDifference("112:1-10", "X-1:X"),
    //    //                            new BibleBookDifference("113:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("114:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("115:1-18", "113:X+8"),
    //    //                            new BibleBookDifference("116:1-9", "114:X"),
    //    //                            new BibleBookDifference("116:10-19", "115:X-9"),
    //    //                            new BibleBookDifference("117:1-2", "X-1:X"),
    //    //                            new BibleBookDifference("118:1-29", "X-1:X"),
    //    //                            new BibleBookDifference("119:1-176", "X-1:X"),
    //    //                            new BibleBookDifference("120:1-7", "X-1:X"),
    //    //                            new BibleBookDifference("121:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("122:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("123:1-4", "X-1:X"),
    //    //                            new BibleBookDifference("124:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("125:1-5", "X-1:X"),
    //    //                            new BibleBookDifference("126:1-6", "X-1:X"),
    //    //                            new BibleBookDifference("127:1-5", "X-1:X"),
    //    //                            new BibleBookDifference("128:1-6", "X-1:X"),
    //    //                            new BibleBookDifference("129:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("130:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("131:1-3", "X-1:X"),
    //    //                            new BibleBookDifference("132:1-18", "X-1:X"),
    //    //                            new BibleBookDifference("133:1-3", "X-1:X"),
    //    //                            new BibleBookDifference("134:1-3", "X-1:X"),
    //    //                            new BibleBookDifference("135:1-21", "X-1:X"),
    //    //                            new BibleBookDifference("136:1-26", "X-1:X"),
    //    //                            new BibleBookDifference("137:1-9", "X-1:X"),
    //    //                            new BibleBookDifference("138:1-8", "X-1:X"),
    //    //                            new BibleBookDifference("139:1-24", "X-1:X"),
    //    //                            new BibleBookDifference("140:1-13", "X-1:X"),
    //    //                            new BibleBookDifference("141:1-10", "X-1:X"),
    //    //                            new BibleBookDifference("142:1-7", "X-1:X"),
    //    //                            new BibleBookDifference("143:1-12", "X-1:X"),
    //    //                            new BibleBookDifference("144:1-15", "X-1:X"),
    //    //                            new BibleBookDifference("145:1-21", "X-1:X"),
    //    //                            new BibleBookDifference("146:1-10", "X-1:X"),
    //    //                            new BibleBookDifference("147:1-11", "X-1:X"),
    //    //                            new BibleBookDifference("147:12-20", "147:X-11")),
    //    //                new BibleBookDifferences(21,
    //    //                            new BibleBookDifference("5:1", "4:17"),
    //    //                            new BibleBookDifference("5:2-20", "5:X-1")),
    //    //                new BibleBookDifferences(22,
    //    //                            new BibleBookDifference("1:1", string.Empty),
    //    //                            new BibleBookDifference("1:2", "1:1"),
    //    //                            new BibleBookDifference("1:3-17", "1:X-1"),
    //    //                            new BibleBookDifference("6:13", "7:1"),
    //    //                            new BibleBookDifference("7:1-13", "7:X+1")),
    //    //                new BibleBookDifferences(23,
    //    //                            new BibleBookDifference("3:19-20", "3:19"),
    //    //                            new BibleBookDifference("3:21-26", "3:X-1")),
    //    //                new BibleBookDifferences(27,
    //    //                            new BibleBookDifference("4:1-3", "3:31-33"),
    //    //                            new BibleBookDifference("4:4-37", "4:X-3")),
    //    //                new BibleBookDifferences(28,
    //    //                            new BibleBookDifference("13:16", "14:1"),
    //    //                            new BibleBookDifference("14:1-9", "14:X+1")),
    //    //                new BibleBookDifferences(32,
    //    //                            new BibleBookDifference("1:17", "2:1"),
    //    //                            new BibleBookDifference("2:1-10", "2:X+1")),
    //    //                new BibleBookDifferences(44,
    //    //                            new BibleBookDifference("19:40-41", "19:40")),
    //    //                new BibleBookDifferences(45,
    //    //                            new BibleBookDifference("16:25-27", "14:24-26")),
    //    //                new BibleBookDifferences(47,
    //    //                            new BibleBookDifference("11:32-33", "11:32"),
    //    //                            new BibleBookDifference("13:12-13", "13:12"),
    //    //                            new BibleBookDifference("13:14", "13:13"))
    //    //            });
    //    //        }

    //    //        return _rst;
    //    //    }
    //    //}
    //}
}
