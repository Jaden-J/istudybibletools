using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;
using BibleCommon.Helpers;

namespace BibleCommon.Common
{
    public enum NotebookType
    {
        Single,
        Bible,
        BibleComments,
        BibleNotesPages,
        BibleStudy
    }

    public enum SectionGroupType
    {
        Bible,
        BibleComments,
        BibleStudy        
    }

    

    [Serializable]
    [XmlRoot(ElementName="IStudyBibleTools_Module")]
    public class ModuleInfo
    {
        [XmlAttribute]
        public string Version { get; set; }

        [XmlIgnore]
        public string ShortName { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(NotebookInfo), ElementName = "Notebook")]
        public List<NotebookInfo> Notebooks { get; set; }

        [XmlElement]
        public BibleStructureInfo BibleStructure { get; set; }

        public bool UseSingleNotebook()
        {
            return Notebooks.Exists(n => n.Type == NotebookType.Single);
        }

        public NotebookInfo GetNotebook(NotebookType notebookType)
        {
            return Notebooks.First(n => n.Type == notebookType);
        }

        // возвращает книгу Библии с учётом всех сокращений
        public BibleBookInfo GetBibleBook(string s)
        {
            s = s.ToLowerInvariant();

            var result = GetBibleBookByExactMatch(s);

            if (result == null && s.Length > 0 && StringUtils.IsDigit(s[0]))
            {

                s = s.Replace(" ", string.Empty); // чтоб находил "1 John", когда в списке сокращений только "1John"
                result = GetBibleBookByExactMatch(s);
            }

            return result;
        }

        private BibleBookInfo GetBibleBookByExactMatch(string s)
        {
            return BibleStructure.BibleBooks.FirstOrDefault(
                book => book.Abbreviations.Contains(s));
        }
    }

    [Serializable]    
    public class NotebookInfo
    {
        [XmlAttribute]
        public NotebookType Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(SectionGroupInfo), ElementName="SectionGroup")]
        public List<SectionGroupInfo> SectionGroups { get; set; }
    }

    [Serializable]
    public class SectionGroupInfo
    {
        [XmlAttribute]
        public SectionGroupType Type { get; set; }

        [XmlAttribute]
        public string Name { get; set; }        
    }

    [Serializable]
    public class BibleStructureInfo
    {
        [XmlAttribute]
        public string OldTestamentName { get; set; }

        [XmlAttribute]
        public int OldTestamentBooksCount { get; set; }

        [XmlAttribute]
        public string NewTestamentName { get; set; }

        [XmlAttribute]
        public int NewTestamentBooksCount { get; set; }        

        [XmlAttribute]
        public string Alphabet { get; set; }  // символы, встречающиеся в названии книг Библии            

        [XmlElement(typeof(BibleBookInfo), ElementName="BibleBook")]
        public List<BibleBookInfo> BibleBooks { get; set; }
    }

    [Serializable]
    public class BibleBookInfo
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public string SectionName { get; set; }

        [XmlElement(typeof(string), ElementName = "Abbreviation")]
        public List<string> Abbreviations { get; set; }
    }
}