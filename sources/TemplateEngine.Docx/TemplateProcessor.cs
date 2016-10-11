namespace TemplateEngine.Docx
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using System.Xml.Linq;

    using DocumentFormat.OpenXml.Packaging;

    using TemplateEngine.Docx.Errors;
    using TemplateEngine.Docx.Processors;

    public class TemplateProcessor : IDisposable
    {
        private readonly WordDocumentContainer _wordDocument;

        private bool _isNeedToRemoveContentControls;

        // private bool _isNeedToNoticeAboutErrors;
        private RenderOptions _renderOptions;

        public TemplateProcessor(string fileName)
            : this(WordprocessingDocument.Open(fileName, true))
        {
        }

        public TemplateProcessor(Stream stream)
            : this(WordprocessingDocument.Open(stream, true))
        {
        }

        public TemplateProcessor(XDocument templateSource, XDocument stylesPart = null, XDocument numberingPart = null)
        {
            _renderOptions = new RenderOptions();

            // _isNeedToNoticeAboutErrors = true;
            _wordDocument = new WordDocumentContainer(templateSource, stylesPart, numberingPart);
        }

        private TemplateProcessor(WordprocessingDocument wordDocument)
        {
            _wordDocument = new WordDocumentContainer(wordDocument);
            _renderOptions = new RenderOptions();

            // _isNeedToNoticeAboutErrors = true;
        }

        public XDocument Document => _wordDocument.MainDocumentPart;

        public Dictionary<string, XDocument> FooterParts => _wordDocument.FooterParts;

        public Dictionary<string, XDocument> HeaderParts => _wordDocument.HeaderParts;

        public IEnumerable<ImagePart> ImagesPart => _wordDocument.ImagesPart;

        public XDocument NumberingPart => _wordDocument.NumberingPart;

        public XDocument StylesPart => _wordDocument.StylesPart;

        public void Dispose()
        {
            _wordDocument?.Dispose();
        }

        public TemplateProcessor FillContent(Content content)
        {
            var processor =
                new ContentProcessor(new ProcessContext(_wordDocument)).SetRemoveContentControls(
                    _isNeedToRemoveContentControls).SetRenderOptions(_renderOptions);
            
            var processResult = processor.FillContent(Document.Root.Element(W.body), content);

            if (_wordDocument.HasFooters)
            {
                foreach (var footer in _wordDocument.FooterParts.Values)
                {
                    var footerProcessResult = processor.FillContent(footer.Element(W.footer), content);
                    processResult.Merge(footerProcessResult);
                }
            }

            if (_wordDocument.HasHeaders)
            {
                foreach (var header in _wordDocument.HeaderParts.Values)
                {
                    var headerProcessResult = processor.FillContent(header.Element(W.header), content);
                    processResult.Merge(headerProcessResult);
                }
            }

            if (_renderOptions.HighlightMissingContent || _renderOptions.HighlightMissingTags)
            {
                AddErrors(processResult.Errors);
            }

            return this;
        }

        public IList<string> GetTags()
        {
            var processor =
                new ContentProcessor(new ProcessContext(_wordDocument)).SetRemoveContentControls(
                    _isNeedToRemoveContentControls).SetRenderOptions(_renderOptions);

            var processResult = processor.FindAllTags(Document.Root.Element(W.body));

            if (_wordDocument.HasFooters)
            {
                foreach (var footer in _wordDocument.FooterParts.Values)
                {
                    var footerProcessResult = processor.FindAllTags(footer.Element(W.footer));
                    processResult.AddRange(footerProcessResult);
                }
            }

            if (_wordDocument.HasHeaders)
            {
                foreach (var header in _wordDocument.HeaderParts.Values)
                {
                    var headerProcessResult = processor.FindAllTags(header.Element(W.header));
                    processResult.AddRange(headerProcessResult);
                }
            }

            // if (_isNeedToNoticeAboutErrors)
            // AddErrors(processResult.Errors);
            return processResult.Distinct().ToList();
        }

        public void SaveChanges()
        {
            _wordDocument.SaveChanges();
        }

        public TemplateProcessor SetRemoveContentControls(bool isNeedToRemove)
        {
            _isNeedToRemoveContentControls = isNeedToRemove;
            return this;
        }

        // public TemplateProcessor SetNoticeAboutErrors(bool isNeedToNotice)
        // {
        // _isNeedToNoticeAboutErrors = isNeedToNotice;
        // return this;
        // }
        public TemplateProcessor SetRenderOptions(RenderOptions renderOptions)
        {
            _renderOptions = renderOptions;
            return this;
        }

        /// <summary>
        /// Adds a list of errors as red text on yellow at the beginning of the document.
        /// </summary>
        /// <param name="errors">List of errors.</param>
        private void AddErrors(IList<IError> errors)
        {
            if (errors.Any())
            {
                Document.Root.Element(W.body)
                    .AddFirst(
                        errors.Select(
                            s =>
                            new XElement(W.p,
                                new XElement(W.r,
                                new XElement(W.rPr,
                                new XElement(W.color, new XAttribute(W.val, _renderOptions.ErrorColor.ToWordColor())),
                                new XElement(W.sz, new XAttribute(W.val, "28")),
                                new XElement(W.szCs, new XAttribute(W.val, "28")),
                                new XElement(W.highlight, new XAttribute(W.val, _renderOptions.ErrorBackground.ToWordColor()))),
                                new XElement(W.t, s.Message)))));
            }
        }
    }
}