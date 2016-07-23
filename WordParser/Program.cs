using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;

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
            ParserSettings settings = new ParserSettings();
            settings.HeaderDepth = 2;
            Parser parser = new Parser(@"C:\Users\ronnel\Documents\docwithpictures.docx", @"C:\Users\ronnel\Documents\docwithpictures_unzpd", settings);
            parser.Run();
            parser.WriteToXml();
        }
    }

    class ParserSettings
    {
        public int HeaderDepth { get; set; }
    }

    class Parser
    {
        public Parser(string docPath, string extractPath, string xmlPath, ParserSettings settings)
        {
            m_docPath = docPath;
            m_extractPath = extractPath;
            m_xmlPath = xmlPath;
            m_settings = settings;
        }

        public void Run()
        {
            PictureUtils.UnzipWordDocument(m_docPath, m_extractPath);
            // Seek to document title, call CreateHeaderSection( title, 0, 0 );
            DocumentIter iter = new DocumentIter(/* Word App */);
            Paragraph titleParagraph = iter.SeekTitle();

            var title = titleParagraph.GetText();
            m_document = CreateHeaderSection(iter, title, 0, 0);
        }

        public void WriteToXml( string xmlPath )
        {

        }

        // Main recursive loop
        private HeaderSection CreateHeaderSection( DocumentIter iter, string header, int start, int depth )
        {
            // TODO: serialize as we parse
            bool fCollapse = depth > m_settings.HeaderDepth;
            HeaderSection section = new HeaderSection(header, start);

            int currHeaderStyle = iter.GetCurrent().HeaderStyle();

            Paragraph p = iter.Next();
            while( p != null && p.HeaderStyle() != currHeaderStyle )
            {
                while( )
                {
                    Content content;
                    section.AddContent(content);
                }
                p = iter.Next();
            }
             
            return section;
        }

        private HeaderSection m_document; // Header section containing all document content
        private string m_docPath; // location of the document file itself
        private string m_extractPath; // location of the unzipped word file
        private string m_xmlPath;
        ParserSettings m_settings;
    }

    class DocumentIter
    {
        public DocumentIter( /* TODO: pass in Word App */ )
        {

        }
        public Paragraph SeekTitle()
        {
            return null;
        }
        public Paragraph First()
        {
            return null;
        }
        public Paragraph Next()
        {
            return null;
        }
        public Paragraph GetCurrent()
        {
            return null;
        }

        private int m_index = 0; // paragraph index
    }
    
    class Paragraph
    {
        public Paragraph( int index /* TODO: pass in Word App*/)
        {
            m_index = index;
        }

        public int HeaderStyle()
        {
            // 0 == title, 2 == header1, 3 == header2, -1 == not header / else
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

        private int m_index;
    }

    class Content
    {
        protected HeaderSection m_section { get; set; }
        protected int m_start { get; set; } // How many lines into the containing HeaderSection we are
    }

    class HeaderSection : Content
    {
        public static int globalId = 0;

        public HeaderSection( string header, int start )
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
        public void AddContent( Content content )
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

        public Picture( int start, string extractPath )
        {
            m_start = start;
            m_id = Picture.generateId();
            m_path = PictureUtils.GetPathFromImageId(m_id, extractPath);
            if( m_path == "" )
            {
                Console.WriteLine("Invalid image ID: {0}", m_id);
            }
        }

        public static int generateId()
        {
            return globalId++;
        }
        private int m_id { get; set; } // Global id iterated from 1. ID = i maps to a file image_i.jpg in the word document
        private string m_path { get; set; }
    }

    class Text : Content
    {
        private string m_text { get; set; }        
        private int m_count { get; set; }
    }
    
    class PictureUtils
    {
        public static List<string> GetDocumentPictures( string extractPath )
        {
            string mediaPath = Path.Combine(extractPath, "word", "media");

            List<string> imgFiles = new List<string>();
            foreach (var file in Directory.GetFiles(mediaPath))
                imgFiles.Add(file);

            return imgFiles;
        }

        public static void UnzipWordDocument( string docPath, string extractPath )
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

        public static string GetPathFromImageId( int id, string extractPath )
        {
            string mediaPath = Path.Combine(extractPath, "word", "media");

            foreach (var file in Directory.GetFiles(mediaPath))
                if (GetImageIdFromPath(file) == id)
                    return file;
            return "";
        }
    }
}
