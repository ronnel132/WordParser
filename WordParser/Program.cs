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

            // TODO: get the input and output paths from the command line
            string docPath = @"C:\Users\ronnel\Downloads\DiffMonsterWhitepaper.docx";
            string outputPath = @"C:\Users\ronnel\Documents\rawoutput.xml";
            //string docPath = @"\\ppt-svc\user\ansanch\DiffMonster Whitepaper.docx";
            //string outputPath = @"\\ppt-svc\user\ansanch\Hackathon\rawoutput.xml";

            ParserSettings settings = new ParserSettings()
            {
                DocPath = docPath,
                ExtractPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(docPath)),
                OutputPath = outputPath,
                HeaderDepth = 2
            };

            Parser parser = new Parser(settings);
            parser.Run();
            parser.WriteOutputFile(outputPath);
        }
    }

    public class ParserSettings
    {
        public int HeaderDepth { get; set; }
        public string ExtractPath { get; set; }
        public string DocPath { get; set; }
        public string OutputPath { get; set; }
        public Document WordDocument { get; set; }
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

            m_settings.WordDocument = doc;

            try
            {
                // Seek to document title, call CreateHeaderSection( title, 0, 0 );
                DocumentIter iter = new DocumentIter(m_settings);
                ParagraphIter titleParagraph = iter.SeekTitle();

                var title = titleParagraph.GetText();

                m_mainHeaderSection = CreateHeaderSection(iter, title, 0, 0);

                m_presentation = new Presentation(m_mainHeaderSection, m_maxDepthEncountered);
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
            m_presentation.Construct();
            m_presentation.WriteToOutputfile(xmlPath);
        }

        // Main recursive loop
        private HeaderSection CreateHeaderSection(DocumentIter iter, string header, int start, int depth)
        {
            if (depth > m_maxDepthEncountered)
                m_maxDepthEncountered = depth;
            // TODO: serialize as we parse
            bool fCollapse = depth >= m_settings.HeaderDepth;
            HeaderSection section = new HeaderSection(header, iter.CurrentCharPosition, start);

            int currHeaderStyle = iter.GetCurrent().HeaderStyle();

            int paragraphCountWithinSection = 0;
            ParagraphIter p = iter.Next();

            while (p != null && p.HeaderStyle() < currHeaderStyle)
            {
                paragraphCountWithinSection++;

                if (p.IsHeaderSection(m_settings.HeaderDepth) && !fCollapse)
                {
                    Content c = CreateHeaderSection(iter, p.GetText(), paragraphCountWithinSection, depth + 1);
                    section.AddContent(c);
                    p = iter.GetCurrent();
                }
                else
                {
                    Content c = p.Next();
                    while (c != null)
                    {
                        section.AddContent(c);
                        c = p.Next();
                    }
                    p = iter.Next();
                }
            }

            return section;
        }

        private HeaderSection m_mainHeaderSection; // Header section containing all document content
        private ParserSettings m_settings;
        private int m_maxDepthEncountered = -1;
        private Presentation m_presentation;
    }

    public class Presentation
    {
        public Presentation( HeaderSection document, int maxDepth )
        {
            m_doc = document;
            m_maxDepth = maxDepth;
        }

        public void Construct()
        {
            List<Slide> slideLst = new List<Slide>();
            string slideTitle = m_doc.GetHeader();
            slideLst.Add(new Slide(slideTitle, null, null)); /* TODO: add image for first slide */

            bool fNewSlideForHeader1s = m_maxDepth > 1;

            ConstructSection(m_doc, fNewSlideForHeader1s);
        }

        public void WriteToOutputfile( string xmlPath)
        {

        }

        private List<Slide> ConstructSection(HeaderSection section, bool fCreateSectionStartSlides)
        {
            List<Slide> slideLst = new List<Slide>();
            List<Content> docContent = m_doc.GetContent();
            foreach( Content c in docContent )
            {
                if(c.GetType().Equals(typeof(HeaderSection)))
                {
                    HeaderSection subSection = (HeaderSection)c;
                    Slide sectionStartSlide = new Slide(section.GetHeader(), null, null); /* TODO: add images for section start slides */
                    slideLst.Add(sectionStartSlide);

                    if (fCreateSectionStartSlides)
                    {
                        List<Slide> subSlides = ConstructSection(subSection, false);
                        slideLst.AddRange(subSlides);
                    }
                }
                else if(c.GetType().Equals(typeof(Text)))
                {

                }
                else if(c.GetType().Equals(typeof(Picture)))
                {

                }
            }
            return null;
        } 

        private List<Slide> ConstructSlidesFromHeaderSection( HeaderSection section )
        {
            throw new NotImplementedException();
        }

        private HeaderSection m_doc;
        private int m_maxDepth;
    }

    public class Slide
    {
        public Slide(string title, List<string> slideContent, List<string> imagePaths)
        {

        }

        public string ConstructSlideString()
        {
            throw new NotImplementedException();
        }
    }

    public class DocumentIter
    {
        public DocumentIter(ParserSettings settings)
        {
            m_settings = settings;
            this.m_index = 0;
        }

        public ParagraphIter SeekTitle()
        {
            Document document = m_settings.WordDocument;

            for (m_index = 1; m_index < document.Paragraphs.Count; m_index++)
            {
                Style style = (Style)document.Paragraphs[m_index].get_Style();
                if (style.NameLocal == "Title")
                {
                    ParagraphIter iter = new ParagraphIter(m_index, m_settings);
                    return iter;
                }
            }

            // if no title found then return an iterator for the first paragraph
            return new ParagraphIter(m_index, m_settings);
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
            return this.FFinished() ? null : new ParagraphIter(m_index, m_settings);
        }

        public long CurrentCharPosition
        {
            get
            {
                if (GetCurrent() != null)
                    return GetCurrent().GetParagraphStartCharPosition();
                return 0;
            }
        }

        public bool FFinished()
        {
            return m_index > m_settings.WordDocument.Paragraphs.Count;
        }

        private int m_index = 0; // paragraph index
        private ParserSettings m_settings;
    }

    public class ParagraphIter
    {
        public ParagraphIter(int index, ParserSettings settings)
        {
            m_paragraphIndex = index;
            m_document = settings.WordDocument;

            m_items = new List<Content>();
            int textStart = GetParagraphStartCharPosition();
            m_items.Add(new Text(textStart, m_document.Paragraphs[m_paragraphIndex].Range.Text));

            // determine if this paragraph contains anything other than text
            // image -> wdInlineShapePicture, chart -> wdInlineShapeChart, diagram -> wdInlineShapeDiagram, smart art = wdInlineShapeSmartArt
            // Note: the Microsoft.Office.Interop.Word.WdInlineShapeType enum also lists a wdInlineShapeLinkedPicture, we may need to support this eventually
            int i = 0;
            foreach (InlineShape shape in m_document.Paragraphs[m_paragraphIndex].Range.InlineShapes)
            {
                i++;

                if (shape != null && shape.Type == Microsoft.Office.Interop.Word.WdInlineShapeType.wdInlineShapePicture)
                {
                    // TODO: ansanch - not sure if the start param on the picture constructor is supposed to be the 
                    // index of the picture in the InlineShapes collection, if not then this needs to be fixed
                    int start = shape.Range.Start;
                    Picture pic = new Picture(i, start, settings.ExtractPath);
                    m_items.Add(pic);
                }
            }

            m_contentIndex = 0;
        }

        public int GetParagraphStartCharPosition()
        {
            return m_document.Paragraphs[m_paragraphIndex].Range.Start;
        }

        /// <summary>
        /// Moves the iterator onto the next Content object within this paragraph i.e. Text, Image, Chart, etc.
        /// </summary>
        /// <returns>null when there is no more content in the paragraph</returns>
        public Content Next(/* int paragraphCount - TODO: not sure if we need paragraphCount here */)
        {
            if (m_contentIndex < m_items.Count)
            {
                return m_items[m_contentIndex++];
            }

            return null;
        }

        /// <summary>
        /// </summary>
        /// <returns>3 == title, 2 == header1, 1 == header2, 0 == not header / else</returns>
        public int HeaderStyle()
        {
            string styleName = ((Style)m_document.Paragraphs[m_paragraphIndex].get_Style()).NameLocal;

            switch (styleName)
            {
                case "Title": return 3;
                case "Heading 1": return 2;
                case "Heading 2": return 1;
                default: return 0;
            }
        }

        public int WordCount()
        {
            return m_document.Paragraphs[m_paragraphIndex].Range.Words.Count;
        }

        //// Returns how many words into the paragraph each image in this paragraph is
        //public List<int> ImageLocations()
        //{
        //    return new List<int>();
        //}

        public string GetText()
        {
            return m_document.Paragraphs[m_paragraphIndex].Range.Text;
        }

        public bool IsHeaderSection(int maxDepth)
        {
            Style style = (Style)m_document.Paragraphs[m_paragraphIndex].get_Style();

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

        private int m_paragraphIndex;
        private int m_contentIndex;
        private Document m_document;
        private List<Content> m_items;
    }

    public class Content
    {
        public Content(long charPosition)
        {
            this.CharPosition = charPosition;
        }

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

        public string GetHeader()
        {
            return m_header;
        }

        public List<Content> GetContent()
        {
            return m_contents;
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
