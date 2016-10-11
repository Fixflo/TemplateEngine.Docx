namespace TemplateEngine.Docx.Processors
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml.Linq;

    using TemplateEngine.Docx.Errors;

    internal class ContentProcessor
    {
        private readonly List<IProcessor> _processors;

        private bool _isNeedToRemoveContentControls;

        private RenderOptions _renderOptions;

        internal ContentProcessor(ProcessContext context)
        {
            _processors = new List<IProcessor>
            {
                new FieldsProcessor(context),
                new TableProcessor(context),
                new ListProcessor(context),
                new ImagesProcessor(context)
            };
        }

        public ProcessResult FillContent(XElement content, IEnumerable<IContentItem> data)
        {
            switch (_renderOptions.RenderMethod)
            {
                case RenderMethod.ByTags:
                    return FillContentByTags(content, data);
                case RenderMethod.ByContent:
                    return FillContentByContent(content, data);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private ProcessResult FillContentByTags(XElement content, IEnumerable<IContentItem> data)
        {
            var result = ProcessResult.NotHandledResult;
            var processedItems = new List<IContentItem>();
            data = data.ToList();

            //var missingControls = FindMissingControls(content, data);
            //foreach (var missingControl in missingControls)
            //{
            //    result.AddError(new ContentControlNotFoundError(missingControl));
            //}

            var allTags = FindAllTags(content);

            foreach (var tag in allTags)
            {
                if (processedItems.Any(i => i.Name == tag))
                {
                    continue;
                }

                var contentControls = FindContentControls(content, tag).ToList();

                // Need to get error message from processor.
                if (!contentControls.Any())
                {
                    contentControls.Add(null);
                }

                foreach (var xElement in contentControls)
                {
                    // COME BACK TO THIS - NOT USING TABLE CONTROLS
                    //if (contentItems.Any(item => item is TableContent) && xElement != null)
                    //{
                    //    var processTableFieldsResult = ProcessTableFields(data.OfType<FieldContent>(), xElement);
                    //    processedItems.AddRange(processTableFieldsResult.HandledItems);

                    //    result.Merge(processTableFieldsResult);
                    //}

                    foreach (var processor in _processors)
                    {
                        var contentItems = data.Where(e => e.Name == tag);

                        if (!contentItems.Any())
                        {
                            contentItems = new List<IContentItem> { new FieldContent { Name = tag, Value = "MISSING" } };
                        }

                        var processorResult = processor.FillContent(xElement, contentItems);

                        processedItems.AddRange(processorResult.HandledItems);
                        result.Merge(processorResult);
                    }
                }
            }

            return result;
        }

        public ProcessResult FillContentByContent(XElement content, IEnumerable<IContentItem> data)
        {
            var result = ProcessResult.NotHandledResult;
            var processedItems = new List<IContentItem>();
            data = data.ToList();

            //var missingControls = FindMissingControls(content, data);
            //foreach (var missingControl in missingControls)
            //{
            //    result.AddError(new ContentControlNotFoundError(missingControl));
            //}

            foreach (var contentItems in data.GroupBy(d => d.Name))
            {
                if (processedItems.Any(i => i.Name == contentItems.Key))
                {
                    continue;
                }

                var contentControls = FindContentControls(content, contentItems.Key).ToList();

                // Need to get error message from processor.
                if (!contentControls.Any())
                {
                    contentControls.Add(null);
                }

                foreach (var xElement in contentControls)
                {
                    if (contentItems.Any(item => item is TableContent) && xElement != null)
                    {
                        var processTableFieldsResult = ProcessTableFields(data.OfType<FieldContent>(), xElement);
                        processedItems.AddRange(processTableFieldsResult.HandledItems);

                        result.Merge(processTableFieldsResult);
                    }

                    foreach (var processor in _processors)
                    {
                        var processorResult = processor.FillContent(xElement, contentItems);

                        processedItems.AddRange(processorResult.HandledItems);
                        result.Merge(processorResult);
                    }
                }
            }

            return result;
        }

        public ProcessResult FillContent(XElement content, Content data)
        {
            return FillContent(content, data.AsEnumerable());
        }

        public ProcessResult FillContent(XElement content, IContentItem data)
        {
            return FillContent(content, new List<IContentItem> { data });
        }

        public List<string> FindAllTags(XElement content)
        {
            var result = ProcessResult.NotHandledResult;

            // var processedItems = new List<IContentItem>();
            // data = data.ToList();
            var contentControls = FindTags(content).ToList();

            // Need to get error message from processor.
            if (!contentControls.Any())
            {
                contentControls.Add(null);
            }

            // foreach (var xElement in contentControls)
            // {
            // if (contentItems.Any(item => item is TableContent) && xElement != null)
            // {
            // var processTableFieldsResult = ProcessTableFields(data.OfType<FieldContent>(), xElement);
            // processedItems.AddRange(processTableFieldsResult.HandledItems);

            // result.Merge(processTableFieldsResult);
            // }

            // foreach (var processor in _processors)
            // {
            // var processorResult = processor.FillContent(xElement, contentItems);

            // processedItems.AddRange(processorResult.HandledItems);
            // result.Merge(processorResult);
            // }
            // }
            return contentControls;
        }

        public ContentProcessor SetRemoveContentControls(bool isNeedToRemove)
        {
            _isNeedToRemoveContentControls = isNeedToRemove;
            foreach (var processor in _processors)
            {
                processor.SetRemoveContentControls(_isNeedToRemoveContentControls);
            }

            return this;
        }

        public ContentProcessor SetRenderOptions(RenderOptions renderOptions)
        {
            _renderOptions = renderOptions;
            foreach (var processor in _processors)
            {
                processor.SetHighlightOptions(_renderOptions);
            }

            return this;
        }

        private IEnumerable<XElement> FindContentControls(XElement content, string tagName)
        {
            return content

                // top level content controls
                .FirstLevelDescendantsAndSelf(W.sdt)

                // with specified tagName
                .Where(sdt => tagName == sdt.SdtTagName());
        }

        private IEnumerable<IContentItem> FindMissingControls(XElement content, IEnumerable<IContentItem> tags)
        {
            var allTags = tags.ToList();
            var allItems = FindTags(content).ToList();
            return allTags.Where(e => !allItems.Contains(e.Name)).ToList();
        }

        private IEnumerable<string> FindTags(XElement content)
        {
            return content

                // top level content controls
                .FirstLevelDescendantsAndSelf(W.sdt)

                // with specified tagName
                .Where(sdt => sdt.SdtTagName() != string.Empty).Select(e => e.SdtTagName());
        }

        /// <summary>
        /// Processes table data that should not be duplicated
        /// </summary>
        /// <param name="fields">Possible fields</param>
        /// <param name="xElement">Table content control</param>
        /// <returns>List of content items that were processed</returns>
        private ProcessResult ProcessTableFields(IEnumerable<FieldContent> fields, XElement xElement)
        {
            var processResult = ProcessResult.NotHandledResult;
            foreach (var fieldContentControl in fields)
            {
                var innerContentControls = FindContentControls(xElement.Element(W.sdtContent), fieldContentControl.Name);
                foreach (var innerContentControl in innerContentControls)
                {
                    var processor = _processors.OfType<FieldsProcessor>().FirstOrDefault();
                    if (processor != null)
                    {
                        var result = processor.FillContent(innerContentControl, fieldContentControl);
                        processResult.Merge(result);
                    }
                }
            }

            return processResult;
        }
    }
}