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

            // When running from Visual Studio (F5) set the command line parameters on the project settings dialog (right click -> properties), debug tab, Startup options -> Command Line Args
            string docPath = args[0];
            string outputPath = args[1];
            

            ParserSettings settings = new ParserSettings()
            {
                DocPath = docPath,
                ExtractPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(docPath)),
                OutputPath = outputPath,
                HeaderDepth = 2,
                MaxCharsPerSlide = 1000
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
        public int MaxCharsPerSlide { get; set; }
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

                m_presentation = new Presentation(m_mainHeaderSection, m_settings.HeaderDepth, m_settings);
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
        private Presentation m_presentation;
    }

    public class Presentation
    {
        public Presentation( HeaderSection document, int maxDepth, ParserSettings settings )
        {
            m_doc = document;
            m_maxDepth = maxDepth;
            m_settings = settings;
        }

        public void Construct()
        {
            List<Slide> slideLst = new List<Slide>();
            string slideTitle = m_doc.GetHeader();
            slideLst.Add(new Slide(slideTitle, null, null)); /* TODO: add image for first slide */
            slideLst.AddRange(ConstructSection(m_doc));
            m_slides = slideLst;
        }

        public void WriteToOutputfile( string xmlPath)
        {
            // TODO
        }

        private List<Slide> ConstructSection(HeaderSection section)
        {
            List<Slide> slideLst = new List<Slide>();
            List<Content> docContent = section.GetContent();
            string sectionHeader = section.GetHeader();
            foreach( Content c in docContent )
            {
                if(c.GetType().Equals(typeof(HeaderSection)))
                {
                    HeaderSection subSection = (HeaderSection)c;

                    List<Slide> subSlides;
                    if(subSection.FContainsHeaderSections())
                    {
                        /* Create the section start slide */
                        Slide sectionStartSlide = new Slide(subSection.GetHeader(), null, null); /* TODO: add images for section start slides */
                        slideLst.Add(sectionStartSlide);

                        /* Recursively construct content inside this header section */
                        subSlides = ConstructSection(subSection);
                    }
                    else
                    {
                        subSlides = ConstructSlidesForSection(subSection);
                    }
                    slideLst.AddRange(subSlides);
                }
                else if(c.GetType().Equals(typeof(Text)))
                {
                   /* TODO */
                }
                else if(c.GetType().Equals(typeof(Picture)))
                {
                   /* TODO */
                }
            }
            return slideLst;
        } 

        /* Construct for section *only* containing text and pictures */
        private List<Slide> ConstructSlidesForSection( HeaderSection section )
        {
            List<Slide> sldList = new List<Slide>();

            List<Content> docContent = section.GetContent();
            string sectionHeader = section.GetHeader();
            List<Picture> pictureList = new List<Picture>();

            int slideCharCount  = 0;

            Slide curSlide = new Slide();
            curSlide.Title = sectionHeader;
            foreach( Content c in docContent )
            {
                if( slideCharCount >= m_settings.MaxCharsPerSlide)
                {
                    sldList.Add(curSlide);
                    curSlide = new Slide();
                    curSlide.Title = sectionHeader;
                    slideCharCount = 0;
                }
                if(c.GetType().Equals(typeof(Text)))
                {
                    Text text = (Text)c;
                    string prettified = text.GetPrettifiedText();
                    if (!String.IsNullOrEmpty(prettified))
                    {
                        slideCharCount += prettified.Length;
                        // TODO: call "GetSummarizedText"
                        curSlide.SlideText.Add(prettified);
                    }
                }
                else if(c.GetType().Equals(typeof(Picture)))
                {
                    // Add to a picture list to store in the correct slide below 
                    Picture picture = (Picture)c;
                    pictureList.Add(picture);
                }
            }

            if (!sldList.Contains(curSlide))
                sldList.Add(curSlide);

            // Place the image in the correct slade based on it's char position
            int charCount = (int) section.CharPosition;
            foreach( Picture pic in pictureList)
            {
                long charPos = pic.CharPosition;
                foreach( Slide slide in sldList )
                {
                    int slideChars = slide.GetCharCount();
                    if( charCount <= charPos && charPos <= charCount + slideChars )
                    {
                        slide.ImagePaths.Add(pic.GetImagePath());
                    }
                    charCount += slideChars;
                }

            }

            return sldList;
        }

        private List<Slide> m_slides;
        private HeaderSection m_doc;
        private int m_maxDepth;
        private ParserSettings m_settings;
    }

    public class Slide
    {
        public Slide()
        {
            SlideText = new List<string>();
            ImagePaths = new List<string>();
        }

        public Slide(string title, List<string> slideText, List<string> imagePaths)
        {
            Title = title;
            SlideText = slideText;
            ImagePaths = imagePaths;
        }

        public int GetCharCount()
        {
            int count = 0;
            foreach( string text in SlideText )
            {
                count += text.Length;
            }
            return count;
        }

        public string ConstructSlideString()
        {
            throw new NotImplementedException();
        }

        public string Title { get; set; }
        public List<string> SlideText { get; set; }
        public List<string> ImagePaths { get; set; }
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
            return m_header.Trim(new char[] { '\r', '\n', '/', '\a' });
        }

        public List<Content> GetContent()
        {
            return m_contents;
        }

        public bool FContainsHeaderSections()
        {
            foreach( Content c in m_contents)
            {
                if (c.GetType().Equals(typeof(HeaderSection)))
                    return true;
            }
            return false;
        }

        public override string ToString()
        {
            return string.Format("HeaderSection: {0}", m_header);
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

        public string GetImagePath()
        {
            return m_path;
        }

        public override string ToString()
        {
            return string.Format("Picture: {0}", Path.GetFileName(m_path));
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

        public string GetText()
        {
            return m_text;
        }

        // Remove newlines, forward slashes, 
        public string GetPrettifiedText()
        {
            return m_text.Trim(new char[] { '\r', '\n', '/' });
        }

        public string GetSummarizedText()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return string.Format("Text: {0}", m_text);
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
