using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordGeneration
{
    class DocGenerationEventArgs : EventArgs
    {
        /// <summary>
        /// Number of documents
        /// </summary>
        public long NumberOfDocs { get; set; }

        /// <summary>
        /// Number of generated docs
        /// </summary>
        public long DocsGenerated { get; set; }
    }
}
