using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WordGeneration
{
    /// <summary>
    /// Class used to generate n word documents
    /// Its role is only to call InternalGeneration asynchronously and 
    /// </summary>
    class DocGeneration
    {
        /// <summary>
        /// Progress of the generation
        /// </summary>
        public event EventHandler<DocGenerationEventArgs> Progress;

        /// <summary>
        /// Generate documents
        /// </summary>
        /// <param name="nbDocs">The number of documents to generate</param>
        /// <param name="destFolder">Destination folder</param>
        /// <param name="nbFilesPerFolder">Number of files per folder</param>
        /// <param name="nbFoldersPerFolder">Number of folders per folder</param>
        /// <param name="nbParaPerDoc">Number of paragraphs per document</param>
        /// <param name="nbWordsPerPara">Number of words per paragraph</param>
        public async Task Generate(long nbDocs, string destFolder, int nbFilesPerFolder = 1000, int nbFoldersPerFolder = 100, int nbParaPerDoc = 10, int nbWordsPerPara = 50)
        {
            var gen = new InternalGeneration(nbDocs, destFolder, nbFilesPerFolder, nbFoldersPerFolder, nbParaPerDoc, nbWordsPerPara);
            gen.Progress += Gen_Progress;
            await Task.Factory.StartNew(() =>
            {
                gen.Launch();
            });
        }

        private void Gen_Progress(object sender, DocGenerationEventArgs e)
        {
            if (Progress != null)
            {
                Progress(this, new DocGenerationEventArgs() { DocsGenerated = e.DocsGenerated, NumberOfDocs = e.NumberOfDocs });
            }
        }
    }
}
