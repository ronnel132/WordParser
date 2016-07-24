using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using Microsoft.Office.Interop.Word;

namespace WordParser
{
    public class Program
    {
        static void Main(string[] args)
        {
            /*
            var imgs = PictureUtils.GetDocumentPictures(@"C:\Users\ronnel\Documents\docwithpictures.docx", @"C:\Users\ronnel\Documents\docwithpictures_unzpd");
            foreach (var img in imgs)
                Console.WriteLine("Path: {0} ImgId: {1}", img, PictureUtils.GetImageIdFromPath(img));
            Console.ReadLine();
            */

            // TODO: get the input file name from the command line
            string docPath = @"\\ppt-svc\user\ansanch\DiffMonster Whitepaper.docx";

            ParserSettings settings = new ParserSettings()
            {
                DocPath = docPath,
                ExtractPath = Path.Combine(Path.GetTempPath(), Path.GetFileName(docPath)),
                OutputPath = @"\\ppt-svc\user\ansanch\Hackathon\rawoutput.xml",
                HeaderDepth = 2
            };

            Parser parser = new Parser(settings);
            parser.Run();
            //parser.WriteToXml();
        }
    }

    public class ParserSettings
    {
        public int HeaderDepth { get; set; }
        public string ExtractPath { get; set; }
        public string DocPath { get; set; }
        public string OutputPath { get; set; }
    }

    public class Parser
    {
        public Parser(ParserSettings settings)
        {
            m_settings = settings;
        }

        public void Run()
        {
            PictureUtils.UnzipWordDocument(m_settings.DocPath, m_settings.ExtractPath);

            // Define an object to pass to the API for missing parameters
            object missing = System.Type.Missing;
            object readOnly = true;
            object docPath = m_settings.DocPath;

            Application word = new Application();
            word.Visible = true; // TODO: remove this for the demo
            Document doc = word.Documents.Open(ref docPath,
                    ref missing, ref readOnly, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing);

            try
            {
                // Seek to document title, call CreateHeaderSection( title, 0, 0 );
                DocumentIter iter = new DocumentIter(doc);
                ParagraphIter titleParagraph = iter.SeekTitle();

                var title = titleParagraph.GetText();

                m_document = CreateHeaderSection(iter, title, 0, 0);
            }
            finally
            {
                // close the document and msword
                ((Microsoft.Office.Interop.Word._Document)doc).Close();
                ((Microsoft.Office.Interop.Word._Application)word).Quit();
            }
        }

        public void WriteOutputFile(string xmlPath)
        {

        }

        // Main recursive loop
        private HeaderSection CreateHeaderSection(DocumentIter iter, string header, int start, int depth)
        {
            // TODO: serialize as we parse
            bool fCollapse = depth >= m_settings.HeaderDepth;
            HeaderSection section = new HeaderSection(header, iter.CurrentCharPosition, start);

            int currHeaderStyle = iter.GetCurrent().HeaderStyle();

            int paragraphCountWithinSection = 0;
            ParagraphIter p = iter.Next();

            while (p != null && p.HeaderStyle() <= currHeaderStyle)
            {
                paragraphCountWithinSection++;

                if (p.IsHeaderSection(m_settings.HeaderDepth) && !fCollapse)
                {
                    Content c = CreateHeaderSection(iter, p.GetText(), paragraphCountWithinSection, depth + 1);
                    section.AddContent(c);
                }
                else
                {
                    Content c = p.Next(/*paragraphCountWithinSection //got confused with this so I removed it for now*/);
                    while (c != null)
                    {
                        section.AddContent(c);
                        c = p.Next(/*paragraphCountWithinSection*/);
                    }
                }

                p = iter.Next();
            }

            return section;
        }

        private HeaderSection m_document; // Header section containing all document content
        private ParserSettings m_settings;

    }

    public class DocumentIter
    {
        public DocumentIter(Document doc)
        {
            m_document = doc;
            this.m_index = 0;
        }
        public ParagraphIter SeekTitle()
        {
            for (m_index = 1; m_index < m_document.Paragraphs.Count; m_index++)
            {
                Style style = (Style)m_document.Paragraphs[m_index].get_Style();
                if (style.NameLocal == "Title")
                {
                    ParagraphIter iter = new ParagraphIter(m_index, m_document);
                    return iter;
                }
            }

            // if no title found then return an iterator for the first paragraph
            return new ParagraphIter(m_index, m_document);
        }

        public ParagraphIter First()
        {
            return null;
        }

        public ParagraphIter Next()
        {
            m_index++;
            return GetCurrent();
        }
        public ParagraphIter GetCurrent()
        {
            return new ParagraphIter(m_index, m_document);
        }

        public long CurrentCharPosition
        {
            get
            {
                // TODO: ansanch - hardcoded to 1 for now
                return 1;
            }
        }

        private int m_index = 0; // paragraph index
        private Document m_document;
    }

    public class ParagraphIter
    {
        public ParagraphIter(int index, Document wordDoc)
        {
            m_index = index;
            m_document = wordDoc;

            var contents = new List<Content>();
            contents.Add(new Text(1, m_document.Paragraphs[m_index].Range.Text));

            // determine if this paragraph contains anything other than text
            // image -> wdInlineShapePicture, chart -> wdInlineShapeChart, diagram -> wdInlineShapeDiagram, smart art = wdInlineShapeSmartArt
            // TODO: the Microsoft.Office.Interop.Word.WdInlineShapeType enum also lists a wdInlineShapeLinkedPicture, I have not tested it yet
            int i = 0;
            foreach (InlineShape shape in m_document.Paragraphs[m_index].Range.InlineShapes)
            {
                i++;

                if (shape != null && shape.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    // TODO: ansanch - not sure if the start param on the picture constructor is supposed to be the 
                    // index of the picture in the InlineShapes collection, if not then this needs to be fixed
                    Picture pic = new Picture(i, 1, "TODO: implement the mapping to the extracted files");
                    contents.Add(pic);
                }
            }

        }

        /// <summary>
        /// Moves the iterator onto the next Content object within this paragraph i.e. Text, Image, Chart, etc.
        /// </summary>
        /// <returns>null when there is no more content in the paragraph</returns>
        public Content Next(/* int paragraphCount - not sure we need a paragraphcount here */)
        {
            return new Text(1, m_document.Paragraphs[m_index].Range.Text);

        }

        /// <summary>
        /// </summary>
        /// <returns>3 == title, 2 == header1, 1 == header2, 0 == not header / else</returns>
        public int HeaderStyle()
        {
            string styleName = ((Style)m_document.Paragraphs[m_index].get_Style()).NameLocal;

            switch (styleName)
            {
                case "Title": return 3;
                case "Header 1": return 2;
                case "Header 2": return 1;
                default: return 0;
            }
        }

        public int WordCount()
        {
            return m_document.Paragraphs[m_index].Range.Words.Count;
        }

        //// Returns how many words into the paragraph each image in this paragraph is
        //public List<int> ImageLocations()
        //{
        //    return new List<int>();
        //}

        public string GetText()
        {
            return m_document.Paragraphs[m_index].Range.Text;
        }

        public bool IsHeaderSection(int maxDepth)
        {
            Style style = (Style)m_document.Paragraphs[m_index].get_Style();

            var headerSectionStyles = new List<string>();
            headerSectionStyles.Add("Title");
            for (int i = 1; i <= maxDepth; i++)
                headerSectionStyles.Add(string.Format("Heading {0}", i));

            if (headerSectionStyles.Contains(style.NameLocal))
            {
                return true;
            }

            return false;
        }

        private int m_index;
        private Document m_document;
    }

    public class Content
    {
        public Content(long charPosition)
        {
            this.CharPosition = CharPosition;
        }

        // TODO: ansanch - do we actually need this?
        //public HeaderSection m_section { get; set; }

        public long CharPosition { get; set; }

        protected int m_start; // How many lines into the containing HeaderSection we are
    }

    public class HeaderSection : Content
    {
        public static int globalId = 0;

        public HeaderSection(string header, long charPosition, int start) : base(charPosition)
        {
            m_header = header;
            m_start = start;
            m_id = HeaderSection.generateId();
            m_contents = new List<Content>();
        }
        public static int generateId()
        {
            return globalId++;
        }
        public void AddContent(Content content)
        {
            m_contents.Add(content);
        }

        private int m_id; // A unique ID for the header
        private List<Content> m_contents; // The list of objects contained in this header
        private string m_header; // The actual title of the header itself
    }

    public class Picture : Content
    {
        public static int globalId = 1;

        public Picture(int start, int charPosition, string extractPath) : base(charPosition)
        {
            m_start = start;
            m_id = Picture.generateId();
            m_path = PictureUtils.GetPathFromImageId(m_id, extractPath);
            if (m_path == "")
            {
                Console.WriteLine("Invalid image ID: {0}", m_id);
            }
        }

        public static int generateId()
        {
            return globalId++;
        }

        private int m_id; // Global id iterated from 1. ID = i maps to a file image_i.jpg in the word document
        private string m_path;
    }

    public class Text : Content
    {
        public Text(long charPosition, string text) : base(charPosition)
        {
            m_text = text;
        }

        private string m_text;
        private int m_count;
    }

    #region utils

    public class PictureUtils
    {
        public static List<string> GetDocumentPictures(string extractPath)
        {
            string mediaPath = Path.Combine(extractPath, "word", "media");

            List<string> imgFiles = new List<string>();
            foreach (var file in Directory.GetFiles(mediaPath))
                imgFiles.Add(file);

            return imgFiles;
        }

        public static void UnzipWordDocument(string docPath, string extractPath)
        {
            if (!Directory.Exists(extractPath))
            {
                string docName = Path.GetFileNameWithoutExtension(docPath);
                string docDirectory = Path.GetDirectoryName(docPath);
                string zipPath = Path.Combine(docDirectory, docName + ".zip");
                File.Copy(docPath, zipPath);
                ZipFile.ExtractToDirectory(zipPath, extractPath);
                File.Delete(zipPath);
            }
        }

        public static int GetImageIdFromPath(string imgPath)
        {
            string imgName = Path.GetFileNameWithoutExtension(imgPath);
            string strId = imgName.Replace("image", "");
            return Int32.Parse(strId);
        }

        public static string GetPathFromImageId(int id, string extractPath)
        {
            string mediaPath = Path.Combine(extractPath, "word", "media");

            foreach (var file in Directory.GetFiles(mediaPath))
                if (GetImageIdFromPath(file) == id)
                    return file;
            return "";
        }
    }

    #endregion

}
