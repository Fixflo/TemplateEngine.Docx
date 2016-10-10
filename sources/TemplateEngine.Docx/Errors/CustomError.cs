namespace TemplateEngine.Docx.Errors
{
    using System;

    internal class CustomError : IError, IEquatable<CustomError>
    {
        internal CustomError(string customMessage)
        {
            Message = customMessage;
        }

        public string Message { get; }

        public bool Equals(IError other)
        {
            if (!(other is CustomError))
            {
                return false;
            }

            return Equals((CustomError)other);
        }

        public bool Equals(CustomError other)
        {
            return Message.Equals(other.Message);
        }

        public override int GetHashCode()
        {
            return Message.GetHashCode();
        }
    }
}