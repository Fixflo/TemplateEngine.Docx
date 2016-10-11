namespace TemplateEngine.Docx.Processors
{
    using System.Collections.Generic;
    using System.Xml.Linq;

    internal interface IProcessor
    {
        ProcessResult FillContent(XElement contentControl, IEnumerable<IContentItem> items);

        IProcessor SetRemoveContentControls(bool isNeedToRemove);

        IProcessor SetHighlightOptions(RenderOptions highlightOptions);

        ProcessResult FillMissingContent(XElement xElement, string name);
    }
}