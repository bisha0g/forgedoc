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
    private readonly Dictionary<string, List<Dictionary<string, string>>> _tablePlaceholders;

    public WordTemplateProcessor(string templatePath)
    {
        if (!File.Exists(templatePath))
            throw new FileNotFoundException("Template file not found", templatePath);
            
        _templatePath = templatePath;
        _placeholders = new Dictionary<string, string>();
        _listPlaceholders = new Dictionary<string, List<Dictionary<string, string>>>();
        _tablePlaceholders = new Dictionary<string, List<Dictionary<string, string>>>();
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

    public WordTemplateProcessor SetTable(string key, List<Dictionary<string, string>> tableData)
    {
        _tablePlaceholders[key] = tableData;
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

            foreach (var tablePlaceholder in _tablePlaceholders)
            {
                ReplaceTable(body, tablePlaceholder.Key, tablePlaceholder.Value);
            }

            wordDoc.MainDocumentPart.Document.Save();
        }
    }

    private void ReplaceText(OpenXmlElement element, string placeholder, string replacement)
    {
        // Get all text elements
        var textElements = element.Descendants<Text>().ToList();
        
        for (int i = 0; i < textElements.Count; i++)
        {
            var currentText = textElements[i].Text;
            
            // Check if this text element contains the start of a placeholder
            if (currentText.Contains("{{"))
            {
                // Build the complete text across multiple elements if needed
                var combinedText = currentText;
                var elementsToRemove = new List<Text>();
                var j = i + 1;
                
                while (j < textElements.Count && !combinedText.Contains("}}"))
                {
                    combinedText += textElements[j].Text;
                    elementsToRemove.Add(textElements[j]);
                    j++;
                }

                // If we found a complete placeholder
                if (combinedText.Contains(placeholder))
                {
                    // Replace the placeholder in the combined text
                    combinedText = combinedText.Replace(placeholder, replacement);
                    
                    // Update the first text element with the entire replaced text
                    textElements[i].Text = combinedText;
                    
                    // Remove the other elements that were part of the split placeholder
                    foreach (var textElement in elementsToRemove)
                    {
                        textElement.Remove();
                    }
                    
                    // Skip the elements we just processed
                    i = j - 1;
                }
            }
            else if (currentText.Contains(placeholder))
            {
                // Handle the simple case where the placeholder is contained within a single text element
                textElements[i].Text = currentText.Replace(placeholder, replacement);
            }
        }
    }

    private void ReplaceLoop(Body body, string listVariable, List<Dictionary<string, string>> values)
    {
        throw new NotImplementedException();
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

    private void ReplaceTable(Body body, string tableVariable, List<Dictionary<string, string>> tableData)
    {
        var paragraphs = body.Descendants<Paragraph>().ToList();
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var paragraphText = paragraphs[i].InnerText;
            if (paragraphText.Contains($"{{{{#{tableVariable}}}}}"))
            {
                // Find the associated table by looking at subsequent elements
                var currentElement = paragraphs[i];
                Table? table = null;
                
                // First try to find table as a sibling
                var nextElement = currentElement.NextSibling();
                while (nextElement != null && table == null)
                {
                    if (nextElement is Table foundTable)
                    {
                        table = foundTable;
                    }
                    nextElement = nextElement.NextSibling();
                }
                
                // If no table found as sibling, search in subsequent elements
                if (table == null)
                {
                    var subsequentElements = currentElement.ElementsAfter();
                    foreach (var element in subsequentElements)
                    {
                        if (element is Table foundTable)
                        {
                            table = foundTable;
                            break;
                        }
                    }
                }

                if (table != null)
                {
                    // Remove the table marker paragraph
                    paragraphs[i].Remove();
                    
                    // Get the template row and its properties
                    var templateRow = table.Elements<TableRow>().FirstOrDefault();
                    if (templateRow != null)
                    {
                        // Store the template cells for formatting reference
                        var templateCells = templateRow.Elements<TableCell>().ToList();
                        
                        // Remove the template row
                        table.RemoveChild(templateRow);

                        // Create rows for each data item
                        foreach (var rowData in tableData)
                        {
                            var newRow = new TableRow();
                            
                            // Use template cells as reference for formatting
                            for (int cellIndex = 0; cellIndex < templateCells.Count && cellIndex < rowData.Count; cellIndex++)
                            {
                                var templateCell = templateCells[cellIndex];
                                var cellValue = rowData.ElementAt(cellIndex).Value;

                                // Clone the template cell to preserve formatting
                                var newCell = (TableCell)templateCell.CloneNode(true);
                                
                                // Clear existing paragraphs in the cloned cell
                                newCell.RemoveAllChildren<Paragraph>();
                                
                                // Add new paragraph with the data
                                var paragraph = new Paragraph(
                                    new Run(
                                        new Text(cellValue)
                                    )
                                );
                                
                                newCell.Append(paragraph);
                                newRow.Append(newCell);
                            }
                            
                            table.Append(newRow);
                        }
                    }
                }
            }
        }
    }
}