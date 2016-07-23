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
    class Program
    {
        static void Main(string[] args)
        {
            /*
            var imgs = PictureUtils.GetDocumentPictures(@"C:\Users\ronnel\Documents\docwithpictures.docx", @"C:\Users\ronnel\Documents\docwithpictures_unzpd");
            foreach (var img in imgs)
                Console.WriteLine("Path: {0} ImgId: {1}", img, PictureUtils.GetImageIdFromPath(img));
            Console.ReadLine();
            */

            Application word = new Application();
            Document doc = new Document();

            ParserSettings settings = new ParserSettings()
            {
                DocPath = @"C:\Users\ronnel\Documents\docwithpictures.docx",
                ExtractPath = @"C:\Users\ronnel\Documents\docwithpictures_unzpd",
                OutputPath = @"\\ppt-svc\user\ansanch\Hackathon\rawoutput.xml",
                WordDoc = doc,
                HeaderDepth = 2
            };

            Parser parser = new Parser(settings);
            parser.Run();
            parser.WriteToXml();
        }
    }

    class ParserSettings
    {
        public int HeaderDepth { get; set; }
        public string ExtractPath { get; set; }
        public string DocPath { get; set; }
        public string OutputPath { get; set; }
        public Document WordDoc { get; set; }
    }

    class Parser
    {
        public Parser(ParserSettings settings)
        {
            m_settings = settings;
        }

        public void Run()
        {
            PictureUtils.UnzipWordDocument(m_settings.DocPath, m_settings.ExtractPath);
            // Seek to document title, call CreateHeaderSection( title, 0, 0 );
            DocumentIter iter = new DocumentIter(m_settings.WordDoc);
            ParagraphIter titleParagraph = iter.SeekTitle();

            var title = titleParagraph.GetText();

            m_document = CreateHeaderSection(iter, title, 0, 0);
        }

        public void WriteOutputFile(string xmlPath)
        {

        }

        // Main recursive loop
        private HeaderSection CreateHeaderSection(DocumentIter iter, string header, int start, int depth)
        {
            // TODO: serialize as we parse
            bool fCollapse = depth >= m_settings.HeaderDepth;
            HeaderSection section = new HeaderSection(header, start);

            int currHeaderStyle = iter.GetCurrent().HeaderStyle();

            ParagraphIter p = iter.Next();
            int paragraphCountWithinSection = 1;

            while (p != null && p.HeaderStyle() <= currHeaderStyle)
            {
                if( p.IsHeaderSection && !fCollapse )
                {
                    CreateHeaderSection(iter, p.GetText(), paragraphCountWithinSection, m_settings.HeaderDepth);
                }
                else
                {
                    Content c = p.Next(paragraphCountWithinSection);
                    while (c != null)
                    {
                        section.AddContent(c);
                        c = p.Next(paragraphCountWithinSection);
                    }
                }

                p = iter.Next();
                paragraphCountWithinSection++;
            }

            return section;
        }

        private HeaderSection m_document; // Header section containing all document content
        private ParserSettings m_settings;

    }

    class DocumentIter
    {
        public DocumentIter(Document doc)
        {

        }
        public ParagraphIter SeekTitle()
        {

            return null;
        }
        public ParagraphIter First()
        {
            return null;
        }
        public ParagraphIter Next()
        {
            return null;
        }
        public ParagraphIter GetCurrent()
        {
            return null;
        }

        private int m_index = 0; // paragraph index
    }

    class ParagraphIter
    {
        public ParagraphIter(int index /* TODO: pass in Word App*/)
        {
            m_index = index;
        }

        public Content Next(int paragraphCount)
        {
            return null;
        }

        /// <summary>
        /// </summary>

        /// <returns>3 == title, 2 == header1, 1 == header2, 0 == not header / else</returns>
        public int HeaderStyle()
        {
            return -1;
        }

        public int WordCount()
        {
            return 0;
        }

        // Returns how many words into the paragraph each image in this paragraph is
        public List<int> ImageLocations()
        {
            return new List<int>();
        }

        public string GetText()
        {
            return null;
        }

        public bool IsHeaderSection
        {
            get
            {
                return false; // TODO: Not implemented yet
            }
        }

        private int m_index;
    }

    class Content
    {
        protected HeaderSection m_section { get; set; }

        protected int m_start; // How many lines into the containing HeaderSection we are
    }

    class HeaderSection : Content
    {
        public static int globalId = 0;

        public HeaderSection(string header, int start)
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

    class Picture : Content
    {
        public static int globalId = 1;

        public Picture(int start, string extractPath)
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

    class Text : Content
    {
        private string m_text;
        private int m_count;
    }

    class PictureUtils
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
}
