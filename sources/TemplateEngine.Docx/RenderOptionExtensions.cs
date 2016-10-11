using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TemplateEngine.Docx
{
    using System.Drawing;

    public static  class RenderOptionExtensions
    {
        public static string ToWordColor(this Color color)
        {
            return color.Name.ToLowerInvariant();
            //return "#" + (color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2")).ToLowerInvariant();
        }
    }
}
