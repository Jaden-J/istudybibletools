using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

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

        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement(typeof(NotebookInfo), ElementName = "Notebook")]
        public List<NotebookInfo> Notebooks { get; set; }

        [XmlElement]
        public BibleStructureInfo BibleStructure { get; set; }

        public NotebookInfo GetNotebook(NotebookType notebookType)
        {
            return Notebooks.First(n => n.Type == notebookType);
        }

        // возвращает книгу Библии с учётом всех сокращений
        public BibleBookInfo GetBibleBook(string s)
        {
            s = s.ToLowerInvariant();

            return BibleStructure.BibleBooks.FirstOrDefault(
                book => book.Shortenings.Contains(s));
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
        public string NewTestamentName { get; set; }        

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

        [XmlElement(typeof(string), ElementName = "Shortening")]
        public List<string> Shortenings { get; set; }
    }
}