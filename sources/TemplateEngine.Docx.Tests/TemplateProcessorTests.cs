﻿using System.Collections.Generic;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TemplateEngine.Docx.TemplateCustomContent;
using TemplateEngine.Docx.Tests.Properties;

namespace TemplateEngine.Docx.Tests
{
    [TestClass]
    public class TemplateProcessorTests
    {
        [TestMethod]
        public void FillingOneTableWithTwoRows()
        {
            var templateDocument = XDocument.Parse(Resources.TemplateWithSingleTable);
            var expectedDocument = XDocument.Parse(Resources.DocumentWithSingleTableFilledWithTwoRows);

            var valuesToFill = new Content
            {
                Tables = new List<TableContent>
                {
                    new TableContent 
                    {
                        Name = "Team Members",
                        Rows = new List<TableRowContent>
                        {
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Eric" },
                                        new FieldContent { Name = "Title", Value = "Program Manager" }
                                    }
                            },
                            new TableRowContent
                            {
                                Fields = new List<FieldContent>
                                    {
                                        new FieldContent { Name = "Name", Value = "Bob" },
                                        new FieldContent { Name = "Title", Value = "Developer" }
                                    }
                            },
                        }
                    }
                }
            };

            var template = new TemplateProcessor(templateDocument)
                .FillContent(valuesToFill);

            var documentXml = template.Document.ToString();

            Assert.AreEqual(expectedDocument.Document.ToString(), documentXml);
        }
    }
}