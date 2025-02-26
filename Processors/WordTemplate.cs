using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ForgeDoc.Processors;

public class WordTemplateProcessor
{
    private readonly string _templatePath;
    private readonly Dictionary<string, string> _placeholders;
    private readonly Dictionary<string, List<Dictionary<string, string>>> _listPlaceholders;

    public WordTemplateProcessor(string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);
            
        _templatePath = templatePath;
        _placeholders = new Dictionary<string, string>();
        _listPlaceholders = new Dictionary<string, List<Dictionary<string, string>>>();
    }

    public static WordTemplateProcessor LoadTemplate(string templatePath)
    {
        return new WordTemplateProcessor(templatePath);
    }

    public WordTemplateProcessor Set(string key, string value)
    {
        _placeholders[key] = value;
        return this;
    }

    public WordTemplateProcessor SetList(string key, List<Dictionary<string, string>> values)
    {
        _listPlaceholders[key] = values;
        return this;
    }

    public void Render(string outputPath)
    {
        File.Copy(_templatePath, outputPath, true);

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(outputPath, true))
        {
            var mainPart = wordDoc.MainDocumentPart ?? throw new InvalidOperationException("Main document part is missing");
            var document = mainPart.Document ?? throw new InvalidOperationException("Document is missing");
            var body = document.Body ?? throw new InvalidOperationException("Document body is missing");

            foreach (var placeholder in _placeholders)
            {
                ReplaceText(body, $"{{{{{placeholder.Key}}}}}", placeholder.Value);
            }

            foreach (var listPlaceholder in _listPlaceholders)
            {
                ReplaceLoop(body, listPlaceholder.Key, listPlaceholder.Value);
            }

            wordDoc.MainDocumentPart.Document.Save();
        }
    }

    private void ReplaceText(OpenXmlElement element, string placeholder, string replacement)
    {
        foreach (var text in element.Descendants<Text>())
        {
            if (text.Text.Contains(placeholder))
            {
                text.Text = text.Text.Replace(placeholder, replacement);
            }
        }
    }

    private void ReplaceLoop(Body body, string listVariable, List<Dictionary<string, string>> values)
    {
        var paragraphs = body.Descendants<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var paragraphText = paragraphs[i].InnerText;
            if (paragraphText.Contains("{%for") && paragraphText.Contains("in"))
            {
                var parts = paragraphText.Split(new[] { "{%for ", " in ", "%}" }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 2)
                {
                    string itemVariable = parts[0].Trim();
                    string listName = parts[1].Trim();

                    if (listName == listVariable)
                    {
                        var startIndex = i;
                        while (i < paragraphs.Count && !paragraphs[i].InnerText.Contains("{%endfor%}"))
                        {
                            i++;
                        }
                        var endIndex = i;
                        var loopTemplate = paragraphs.Skip(startIndex + 1).Take(endIndex - startIndex - 1).ToList();

                        var parent = paragraphs[startIndex].Parent;
                        for (int j = startIndex; j <= endIndex; j++)
                        {
                            parent.RemoveChild(paragraphs[j]);
                        }

                        foreach (var item in values)
                        {
                            foreach (var templateParagraph in loopTemplate)
                            {
                                var newParagraph = (Paragraph)templateParagraph.CloneNode(true);
                                foreach (var kvp in item)
                                {
                                    ReplaceText(newParagraph, $"{{{{{itemVariable}.{kvp.Key}}}}}", kvp.Value);
                                }
                                parent.AppendChild(newParagraph);
                            }
                        }
                    }
                }
            }
        }
    }
}