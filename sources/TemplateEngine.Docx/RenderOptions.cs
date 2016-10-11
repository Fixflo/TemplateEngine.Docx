namespace TemplateEngine.Docx
{
    public enum RenderMethod
    {
        ByTags,
        ByContent
    }

    public class RenderOptions
    {
        public RenderOptions()
        {
            Color = "red";
            Background = "yellow";
            MissingContentText = "MISSING";
            RenderMethod = RenderMethod.ByContent;
            HighlightMissingContent = true;
            HighlightMissingTags = true;
        }

        public string Color { get; set; }

        public string Background { get; set; }

        public RenderMethod RenderMethod { get; set; }

        public bool HighlightMissingTags { get; set; }

        public bool HighlightMissingContent { get; set; }

        public string MissingContentText { get; set; }
    }
}