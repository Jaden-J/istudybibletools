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