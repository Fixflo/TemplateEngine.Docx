namespace TemplateEngine.Docx
{
    using System.Drawing;

    public enum RenderMethod
    {
        ByTags, 

        ByContent
    }

    public class RenderOptions
    {
        public RenderOptions()
        {
            HighlightColor = Color.White;
            HighlightBackground = Color.Magenta;
            MissingContentText = "[{0}] VALUE MISSING";
            RenderMethod = RenderMethod.ByContent;
            HighlightMissingContent = true;
            HighlightMissingTags = true;
            ErrorColor = Color.Red;
            ErrorBackground = Color.Yellow;
        }

        public Color ErrorBackground { get; set; }

        public Color ErrorColor { get; set; }

        public Color HighlightBackground { get; set; }

        public Color HighlightColor { get; set; }

        public bool HighlightMissingContent { get; set; }

        public bool HighlightMissingTags { get; set; }

        public bool HighlightTags { get; set; }

        public string MissingContentText { get; set; }

        public RenderMethod RenderMethod { get; set; }
    }
}