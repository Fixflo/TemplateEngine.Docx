namespace TemplateEngine.Docx.Errors
{
    using System;

    internal class ContentControlNotFoundError : IError, IEquatable<ContentControlNotFoundError>
    {
        private const string ErrorMessageTemplate = "{0} Content Control '{1}' not found.";

        internal ContentControlNotFoundError(IContentItem contentItem)
        {
            ContentItem = contentItem;
        }

        public IContentItem ContentItem { get; }

        public string Message => string.Format(ErrorMessageTemplate, ContentItem.GetContentItemName(), ContentItem.Name);

        public bool Equals(ContentControlNotFoundError other)
        {
            return other != null && other.ContentItem.Equals(ContentItem);
        }

        public bool Equals(IError other)
        {
            if (!(other is ContentControlNotFoundError))
            {
                return false;
            }

            return Equals((ContentControlNotFoundError)other);
        }

        public override int GetHashCode()
        {
            return ContentItem.GetHashCode();
        }
    }
}