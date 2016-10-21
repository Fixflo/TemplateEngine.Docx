namespace TemplateEngine.Docx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;

    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class TemplateMerger : IDisposable
    {
        private readonly IEnumerable<MemoryStream> _wordDocuments;

        public TemplateMerger(IEnumerable<string> files)
        {
            var s = new List<MemoryStream>(0);
            foreach (var file in files)
            {
                using (var fileStream = File.Open(file, FileMode.Open))
                {
                    var memoryStream = new MemoryStream();
                    fileStream.CopyTo(memoryStream);
                    s.Add(memoryStream);
                }
            }

            _wordDocuments = s;
        }

        public TemplateMerger(IEnumerable<MemoryStream> streams)
        {
            _wordDocuments = streams.ToList();
        }

        public void Dispose()
        {
            _wordDocuments.ToList().ForEach(e => e.Dispose());
        }

        public MemoryStream Merge()
        {
            var firstDocument = _wordDocuments.FirstOrDefault();

            if (firstDocument == null)
            {
                return new MemoryStream();
            }

            using (var myDoc = WordprocessingDocument.Open(firstDocument, true))
            {
                var mainPart = myDoc.MainDocumentPart;

                for (var i = 1; i < _wordDocuments.Count(); i++)
                {
                    // add in a page break
                    var last =
                        myDoc.MainDocumentPart.Document.Body.Elements()
                            .LastOrDefault(e => e is Paragraph || e is AltChunk);
                    last?.InsertAfterSelf(new Paragraph(new Run(new Break { Type = BreakValues.Page })));

                    // get the next document in the list as a stream
                    var stream = _wordDocuments.ToList()[i];
                    stream.Position = 0;

                    // add in each stream as a chunk
                    var altChunkId = $"AltChunkId{i}";

                    var chunk = mainPart.AddAlternativeFormatImportPart(
                        AlternativeFormatImportPartType.WordprocessingML, 
                        altChunkId);

                    chunk.FeedData(stream);

                    var altChunk = new AltChunk { Id = altChunkId };

                    // add in the document as a new chunk
                    mainPart.Document.Body.InsertAfter(altChunk, mainPart.Document.Body.Elements<Paragraph>().Last());

                    // save the document
                    mainPart.Document.Save();
                }
            }

            firstDocument.Position = 0;
            return firstDocument;
        }
    }
}