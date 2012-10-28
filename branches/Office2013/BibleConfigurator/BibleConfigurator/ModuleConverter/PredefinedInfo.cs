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
                        SectionGroups = RussianStrongNotebookBibleSectionGroups
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

        private static List<SectionGroupInfo> RussianStrongNotebookBibleSectionGroups
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
                                Type = ContainerType.OldTestament,
                                StrongPrefix = "H"
                            },
                            new SectionGroupInfo() 
                            { 
                                Name = "Новый Завет", 
                                CheckSectionsCount = true, 
                                SectionsCount = 27, 
                                Type = ContainerType.NewTestament,
                                StrongPrefix = "G"
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
                _rst.Add(80);
                _rst.Add(69);
                _rst.Add(67);
                _rst.Add(201);
                _rst.Add(70);
                _rst.Add(79);
                _rst.Add(71);
                _rst.Add(72);
                _rst.Add(73);
                _rst.Add(77);
                _rst.Add(81);                
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
}
