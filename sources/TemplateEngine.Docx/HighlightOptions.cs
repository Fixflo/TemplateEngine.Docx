namespace TemplateEngine.Docx
{
    public class HighlightOptions
    {
        public HighlightOptions()
        {
            Color = "red";
            Background = "yellow";
        }

        public string Color { get; set; }

        public string Background { get; set; }
    }
}