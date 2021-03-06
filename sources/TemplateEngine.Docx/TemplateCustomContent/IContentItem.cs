﻿using System;

namespace TemplateEngine.Docx
{
	public interface IContentItem : IEquatable<IContentItem>
	{
		string Name { get; set; }

        bool IsMissing { get; set; }
	}
}
