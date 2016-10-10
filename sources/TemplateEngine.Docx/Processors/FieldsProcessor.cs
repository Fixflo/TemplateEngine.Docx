namespace TemplateEngine.Docx.Processors
{
    using System.Collections.Generic;
    using System.Xml.Linq;

    using TemplateEngine.Docx.Errors;

    internal class FieldsProcessor : IProcessor
    {
        private bool _isNeedToRemoveContentControls;

        private HighlightOptions _highlightOptions;

        private readonly ProcessContext _context;
        public FieldsProcessor(ProcessContext context)
        {
            _context = context;
        }

        public ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items)
        {
            var processResult = ProcessResult.NotHandledResult;

            foreach (var contentItem in items)
            {
                processResult.Merge(FillContent(contentControl, contentItem));
            }

            if (processResult.Success && _isNeedToRemoveContentControls)
            {
                contentControl.RemoveContentControl();
            }

            return processResult;
        }

        public ProcessResult FillContent(XElement contentControl, IContentItem item)
        {
            var processResult = ProcessResult.NotHandledResult;
            if (!(item is FieldContent))
            {
                processResult = ProcessResult.NotHandledResult;
                return processResult;
            }

            var field = (FieldContent)item;

            // If there isn't a field with that name, add an error to the error string,
            // and continue with next field.
            if (contentControl == null)
            {
                processResult.AddError(new ContentControlNotFoundError(field));
                return processResult;
            }

            contentControl.ReplaceContentControlWithNewValue(field.Value, _highlightOptions);

            processResult.AddItemToHandled(item);

            return processResult;
        }

        public IProcessor SetRemoveContentControls(bool isNeedToRemove)
        {
            _isNeedToRemoveContentControls = isNeedToRemove;
            return this;
        }

        public IProcessor SetHighlightOptions(HighlightOptions options)
        {
            _highlightOptions = options;
            return this;
        }
    }
}