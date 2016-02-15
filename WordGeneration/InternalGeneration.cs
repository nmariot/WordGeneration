using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace WordGeneration
{
    /// <summary>
    /// Internal generation
    /// </summary>
    class InternalGeneration
    {
        static readonly string[] WORDS = { "a", "ac", "adipiscing", "aenean", "aliquam", "aliquet", "amet", "ante", "arcu", "at", "augue", "bibendum", "blandit", "congue", "consectetuer", "consectetur", "consequat", "convallis", "cras", "cubilia", "curabitur", "curae;", "cursus", "diam", "dignissim", "dolor", "dui", "duis", "egestas", "eleifend", "elementum", "elit", "enim", "erat", "eros", "est", "et", "eu", "euismod", "faucibus", "felis", "fermentum", "feugiat", "fusce", "gravida", "hendrerit", "iaculis", "id", "imperdiet", "in", "integer", "ipsum", "justo", "lacus", "lectus", "leo", "libero", "ligula", "lobortis", "lorem", "luctus", "maecenas", "magna", "massa", "mauris", "metus", "mi", "molestie", "mollis", "morbi", "nec", "nibh", "nisl", "non", "nonummy", "nulla", "nunc", "odio", "orci", "ornare", "pede", "pellentesque", "pharetra", "placerat", "porttitor", "posuere", "praesent", "pretium", "primis", "proin", "pulvinar", "purus", "quam", "quis", "rhoncus", "risus", "rutrum", "sapien", "scelerisque", "sed", "sem", "semper", "sit", "sodales", "sollicitudin", "suscipit", "suspendisse", "tellus", "tempor", "tempus", "tincidunt", "tortor", "tristique", "turpis", "ullamcorper", "ultrices", "ultricies", "ut", "varius", "vehicula", "vel", "velit", "vestibulum", "vitae", "vivamus", "volutpat", "vulputate" };

        /// <summary>
        /// Progress of the generation
        /// </summary>
        public event EventHandler<DocGenerationEventArgs> Progress;

        /// <summary>
        /// Total number of documents
        /// </summary>
        public long NbTotalDocs { get; private set; }

        /// <summary>
        /// Destination folder
        /// </summary>
        public string DestFolder { get; private set; }

        /// <summary>
        /// Number of files per folder
        /// </summary>
        public int NbFilesPerFolder { get; private set; }

        /// <summary>
        /// Number of folders per folder
        /// </summary>
        public int NbFoldersPerFolder { get; private set; }

        /// <summary>
        /// Number of paragraphs per document
        /// </summary>
        public int NbParagraphesPerDoc { get; private set; }

        /// <summary>
        /// Number of words per paragraphs
        /// </summary>
        public int NbWordsPerParagraph { get; private set; }

        private int _depth;

        #region Random management in multithread context       
        private static Random _global = new Random();
        [ThreadStatic]
        private static Random _local;

        /// <summary>
        /// Return a random word
        /// </summary>
        /// <returns>A work chosen into the array</returns>
        private string GetRandomWord()
        {
            Random inst = _local;
            if (inst == null)
            {
                int seed;
                lock (_global) seed = _global.Next();
                _local = inst = new Random(seed);
            }

            return WORDS[inst.Next(0, WORDS.Length - 1)];
        }
        #endregion

        private int _nbDocsGenerated = 0;

        /// <summary>
        /// Get an instance of the Generation class
        /// </summary>
        /// <param name="nbTotalDocs"></param>
        /// <param name="destFolder">Destination folder</param>
        /// <param name="nbFilesPerFolder">Number of files per folder</param>
        /// <param name="nbFoldersPerFolder">Number of folders per folder</param>
        /// <param name="nbParaPerDoc">Number of paragraphs per document</param>
        /// <param name="nbWordsPerPara">Number of words per paragraph</param>
        public InternalGeneration(long nbTotalDocs, string destFolder, int nbFilesPerFolder, int nbFoldersPerFolder, int nbParaPerDoc, int nbWordsPerPara)
        {
            this.NbTotalDocs = nbTotalDocs;
            this.DestFolder = destFolder;
            this.NbFilesPerFolder = nbFilesPerFolder;
            this.NbFoldersPerFolder = nbFoldersPerFolder;
            this.NbParagraphesPerDoc = nbParaPerDoc;
            this.NbWordsPerParagraph = nbWordsPerPara;

            // Calculate the depth of our folders
            for (_depth = 0; ; _depth++)
            {
                if (nbTotalDocs <= this.NbFilesPerFolder * Math.Pow(this.NbFoldersPerFolder, _depth))
                {
                    break;
                }
            }
        }

        /// <summary>
        /// Launch the generation with parameters specified in the constructor
        /// </summary>
        public void Launch()
        {
            _nbDocsGenerated = 0;
            Parallel.For(_nbDocsGenerated, this.NbTotalDocs, GenerateDocument);
        }

        /// <summary>
        /// Generate a Word document
        /// </summary>
        /// <param name="number">Document number</param>
        public void GenerateDocument(long number)
        {
            // Evaluate final path
            string finalPath = this.DestFolder;
            for (int folder = 0; folder < _depth; folder++)
            {
                long n = (long)Math.Pow(this.NbFoldersPerFolder, _depth - folder - 1) * this.NbFoldersPerFolder;
                finalPath = Path.Combine(finalPath, "Folder" + (number / n).ToString());
            }
            Directory.CreateDirectory(finalPath);

            string path = Path.Combine(finalPath, $"{GetRandomWord()} {GetRandomWord()} {GetRandomWord()} {number.ToString()}.docx");

            // Writing in the word document
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(path, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());
                for (var numPara = 0; numPara < this.NbParagraphesPerDoc; numPara++)
                {
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    StringBuilder sb = new StringBuilder();
                    for (var numWord = 0; numWord < this.NbWordsPerParagraph; numWord++)
                    {
                        sb.Append(GetRandomWord() + " ");
                    }
                    run.AppendChild(new Text(sb.ToString()));
                }
            }

            // Raising the progress event
            _nbDocsGenerated++;
            OnProgress(_nbDocsGenerated, this.NbTotalDocs);
        }



        private void OnProgress(long current, long nbDocs)
        {
            if (Progress != null)
            {
                Progress(this, new DocGenerationEventArgs() { DocsGenerated = current, NumberOfDocs = nbDocs });
            }
        }
    }
}
