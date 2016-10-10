namespace TemplateEngine.Docx.Processors
{
    using System.Collections.Generic;

    internal class ProcessContext
    {
        internal ProcessContext(WordDocumentContainer document)
        {
            Document = document;
            LastNumIds = new Dictionary<int, int>();
        }

        internal WordDocumentContainer Document { get; private set; }

        internal Dictionary<int, int> LastNumIds { get; private set; }
    }
}