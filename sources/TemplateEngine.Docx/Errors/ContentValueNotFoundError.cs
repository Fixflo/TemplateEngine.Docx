namespace TemplateEngine.Docx.Errors
{
    using System;

    internal class ContentValueNotFoundError : IError, IEquatable<ContentValueNotFoundError>
    {
        private const string ErrorMessageTemplate = "The value for '{0}' was not found.";

        internal ContentValueNotFoundError(IContentItem contentItem)
        {
            ContentItem = contentItem;
        }

        public IContentItem ContentItem { get; }

        public string Message => string.Format(ErrorMessageTemplate, ContentItem.Name);

        public bool Equals(ContentValueNotFoundError other)
        {
            return other != null && other.ContentItem.Equals(ContentItem);
        }

        public bool Equals(IError other)
        {
            if (!(other is ContentValueNotFoundError))
            {
                return false;
            }

            return Equals((ContentValueNotFoundError)other);
        }

        public override int GetHashCode()
        {
            return ContentItem.GetHashCode();
        }
    }
}