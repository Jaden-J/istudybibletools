using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace BibleConfigurator.ModuleConverter
{
    [XmlRoot("ID")]
    public class BibleBooksInfo
    {
        [XmlAttribute("descr")]
        public string Descr { get; set; }

        [XmlAttribute("alphabet")]
        public string Alphabet { get; set; }

        [XmlAttribute]
        public string ChapterString { get; set; }

        [XmlElement(typeof(BookInfo), ElementName = "BOOK")]
        public List<BookInfo> Books { get; set; }

        public BibleBooksInfo()
        {
            this.Books = new List<BookInfo>();
        }
    }
    
    public class BookInfo
    {
        [XmlAttribute("bnumber")]
        public int Index { get; set; }

        [XmlAttribute("bshort")]
        public string ShortNamesXMLString { get; set; }

        [XmlText]        
        public string Name { get; set; }
    }
}
