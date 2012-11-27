using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using BibleCommon.Common;
using BibleCommon.Scheme;
using BibleCommon.Helpers;

namespace BibleConfigurator.ModuleConverter
{
    public class TextFilesConverter
    {
        public string SourceDirectory { get; set; }
        public string TargetFilePath { get; set; }        
        public XMLBIBLE Bible { get; set; }
        public List<Exception> Errors { get; set; }
        public Encoding Encoding { get; set; }

        public TextFilesConverter(string sourceDirectory, string targetFilePath)
        {
            this.SourceDirectory = sourceDirectory;
            this.TargetFilePath = targetFilePath;

            if (!Directory.Exists(TargetFilePath))
                Directory.CreateDirectory(TargetFilePath);
            
            Bible = new XMLBIBLE();            
            Bible.BIBLEBOOK = new BIBLEBOOK[0];
            Errors = new List<Exception>();

            if (SourceDirectory.EndsWith("t", StringComparison.InvariantCultureIgnoreCase))
                Encoding = Encoding.GetEncoding(950);
            else if (SourceDirectory.EndsWith("s", StringComparison.InvariantCultureIgnoreCase))
                Encoding = Encoding.GetEncoding(936);
            else
                throw new Exception("Encoding can not be determined");
        }        

        public void Convert()
        {
            foreach (var bookDirectory in Directory.GetDirectories(SourceDirectory, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    var bookName = Path.GetFileName(bookDirectory);
                    var vp = new VersePointer(bookName, 1);
                    if (!vp.IsValid)
                        throw new Exception(string.Format("The book '{0}' can not be determined.", bookName));

                    ReadBookContent(vp.Book.Index, bookDirectory);
                }
                catch (Exception ex)
                {
                    Errors.Add(ex);
                }
            }


            Bible.BIBLEBOOK = Bible.Books.OrderBy(b => b.Index).ToArray();
            
            Utils.SaveToXmlFile(Bible, Path.Combine(TargetFilePath, BibleCommon.Consts.Constants.BibleContentFileName));
        }

        private void ReadBookContent(int bookIndex, string bookDirectory)
        {
            try
            {
                var book = new BIBLEBOOK() { bnumber = bookIndex.ToString() };
                book.Items = new CHAPTER[0];
                Bible.BIBLEBOOK = Bible.BIBLEBOOK.Add(book).ToArray();

                foreach (var filePath in Directory.GetFiles(bookDirectory))
                {
                    ReadChapterContent(ref book, filePath);
                }

                if (book.Items == null || book.Items.Length == 0)
                    throw new Exception(string.Format("No chapters found in book {0}", book.Index));

                book.Items = book.Items.OrderBy(c => c.Index).ToArray();
            }
            catch (Exception ex)
            {
                Errors.Add(ex);
            }
        }

        private void ReadChapterContent(ref BIBLEBOOK book, string filePath)
        {
            try
            {
                var chapterIndex = StringUtils.GetStringFirstNumber(Path.GetFileNameWithoutExtension(filePath));
                var chapter = new CHAPTER() { cnumber = chapterIndex.ToString() };
                book.Items = book.Items.Add(chapter).ToArray();

                string fileContent = File.ReadAllText(filePath, Encoding);

                var classTString = "class=\"t\"";
                var classVString = "class=\"v\"";
                var getVerseContentByClass = fileContent.IndexOf(classTString) != -1;
                var verseContentStartSearchString = getVerseContentByClass ? classTString : "<td>";

                var cursor = fileContent.IndexOf(classVString, StringComparison.InvariantCultureIgnoreCase);
                while (cursor > -1)
                {
                    var verseIndexString = GetTagValue(fileContent, cursor);
                    string verseIndex;

                    if (verseIndexString.IndexOf(':') != -1)
                    {
                        var verseIndexParts = verseIndexString.Split(new char[] { ':' });
                        var verseChapterIndex = verseIndexParts[0];
                        if (chapterIndex.ToString() != verseChapterIndex)
                            throw new Exception(string.Format("Wrong chapter index of book {2}: {0} != {1}", chapterIndex, verseChapterIndex, book.Index));

                        verseIndex = verseIndexParts[1];
                    }
                    else
                        verseIndex = verseIndexString;

                    var verseContentStartIndex = fileContent.IndexOf(verseContentStartSearchString, cursor + 1, StringComparison.InvariantCultureIgnoreCase);

                    if (verseContentStartIndex == -1)
                        throw new Exception(string.Format("VerseContent not found in chapter {0} {1}", book.Index, chapterIndex));

                    var versecontent = GetTagValue(fileContent, verseContentStartIndex);                    

                    chapter.Items = chapter.Items.Add(new VERS() { vnumber = verseIndex, Items = new object[] { versecontent } }).ToArray();

                    cursor = fileContent.IndexOf("class=\"v\"", cursor + 1, StringComparison.InvariantCultureIgnoreCase);
                }

                if (chapter.Items == null || chapter.Items.Length == 0)
                    throw new Exception(string.Format("No verses found in chapter {0} {1}", book.Index, chapterIndex));
            }
            catch (Exception ex)
            {
                Errors.Add(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="s"></param>
        /// <param name="index">позиция тэга - до начала символа '>'</param>
        /// <returns></returns>
        private string GetTagValue(string s, int index)
        {
            var startIndex = s.IndexOf('>', index + 1);
            if (startIndex != -1)
            {
                var endIndex = s.IndexOf('<', startIndex + 1);
                if (endIndex != -1)
                {
                    return  StringUtils.GetText(s.Substring(startIndex + 1, endIndex - startIndex - 1).Trim(new char[] { '\r', '\n', '\t', ' ', '　' }));
                }
            }

            return null;
        }
    }
}
