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
        var paragraphs = body.Descendants<Paragraph>().ToList();
        
        // Find loop start and end markers
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var startParagraph = paragraphs[i];
            var startText = startParagraph.InnerText;
            
            // Check if this paragraph contains a loop start marker
            if (startText.Contains($"{{%for ") && startText.Contains($" in {listVariable}%}}"))
            {
                // Extract the item variable name
                int forIndex = startText.IndexOf("{%for ") + 6;
                int inIndex = startText.IndexOf(" in ");
                if (inIndex > forIndex)
                {
                    string itemVariable = startText.Substring(forIndex, inIndex - forIndex).Trim();
                    
                    // Find the end of the loop
                    int endIndex = -1;
                    for (int j = i + 1; j < paragraphs.Count; j++)
                    {
                        if (paragraphs[j].InnerText.Contains("{%endfor%}"))
                        {
                            endIndex = j;
                            break;
                        }
                    }
                    
                    if (endIndex > i)
                    {
                        // Get the parent element that contains the paragraphs
                        var parent = startParagraph.Parent;
                        
                        // Store the template paragraphs (between start and end markers)
                        var templateParagraphs = new List<Paragraph>();
                        for (int j = i + 1; j < endIndex; j++)
                        {
                            templateParagraphs.Add((Paragraph)paragraphs[j].CloneNode(true));
                        }
                        
                        // Remove the original loop paragraphs (including start and end markers)
                        for (int j = endIndex; j >= i; j--)
                        {
                            paragraphs[j].Remove();
                        }
                        
                        // Insert new paragraphs for each item in the list
                        int insertPosition = i;
                        foreach (var item in values)
                        {
                            foreach (var templateParagraph in templateParagraphs)
                            {
                                var newParagraph = (Paragraph)templateParagraph.CloneNode(true);
                                
                                // Replace placeholders in the paragraph
                                foreach (var kvp in item)
                                {
                                    string placeholder = $"{{{{{itemVariable}.{kvp.Key}}}}}";
                                    ReplaceText(newParagraph, placeholder, kvp.Value);
                                }
                                
                                // Insert the new paragraph
                                if (insertPosition == i)
                                {
                                    // Insert at the position of the start marker
                                    parent.InsertAt(newParagraph, insertPosition);
                                }
                                else
                                {
                                    // Insert after the last inserted paragraph
                                    var lastParagraph = parent.Elements<Paragraph>().ElementAt(insertPosition);
                                    parent.InsertAfter(newParagraph, lastParagraph);
                                }
                                insertPosition++;
                            }
                        }
                        
                        // Update the paragraphs list since we've modified the document
                        paragraphs = body.Descendants<Paragraph>().ToList();
                        
                        // Adjust the index to continue processing after the newly inserted paragraphs
                        i = insertPosition - 1;
                    }
                }
            }
        }
    }

    private void ReplaceTable(Body body, string tableVariable, List<Dictionary<string, string>> tableData)
    {
        // First try to find tables with the marker in a paragraph before the table
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
                    
                    // Process the table
                    ProcessTableData(table, tableData);
                    return; // Table found and processed
                }
            }
        }
        
        // If we get here, we didn't find the marker in a paragraph before a table
        // Now look for tables that have the marker inside a cell
        var tables = body.Descendants<Table>().ToList();
        foreach (var table in tables)
        {
            bool markerFound = false;
            
            // Check if any cell contains our marker
            foreach (var cell in table.Descendants<TableCell>())
            {
                if (cell.InnerText.Contains($"{{{{#{tableVariable}}}}}"))
                {
                    markerFound = true;
                    break;
                }
            }
            
            if (markerFound)
            {
                // Process the table
                ProcessTableData(table, tableData);
                break;
            }
        }
    }
    
    private void ProcessTableData(Table table, List<Dictionary<string, string>> tableData)
    {
        // Get all rows in the table
        var rows = table.Elements<TableRow>().ToList();
        
        // We need at least one row as a template
        if (rows.Count == 0)
            return;
            
        // Store the template row (first row)
        var templateRow = rows[0];
        
        // Get the template cells
        var templateCells = templateRow.Elements<TableCell>().ToList();
        
        // Check if we have a table marker in the first cell
        bool hasTableMarker = false;
        string firstCellText = templateCells.FirstOrDefault()?.InnerText ?? string.Empty;
        if (firstCellText.Contains("{{#") && firstCellText.Contains("}}"))
        {
            hasTableMarker = true;
        }
        
        // Remove all existing rows
        while (table.Elements<TableRow>().Any())
        {
            table.RemoveChild(table.Elements<TableRow>().First());
        }
        
        // Create header row with column names
        var headerRow = new TableRow();
        foreach (var key in tableData.FirstOrDefault()?.Keys ?? new Dictionary<string, string>().Keys)
        {
            var headerCell = new TableCell(new Paragraph(new Run(new Text(key))));
            headerRow.AppendChild(headerCell);
        }
        table.AppendChild(headerRow);
        
        // Create rows for each data item
        foreach (var rowData in tableData)
        {
            var newRow = new TableRow();
            
            // Create cells for each key in the dictionary
            foreach (var key in rowData.Keys)
            {
                var cellValue = rowData[key];
                var newCell = new TableCell();
                UpdateCellContent(newCell, cellValue);
                newRow.AppendChild(newCell);
            }
            
            table.AppendChild(newRow);
        }
    }
    
    private void UpdateCellContent(TableCell cell, string newContent)
    {
        // Get the first paragraph or create one if none exists
        var paragraph = cell.Elements<Paragraph>().FirstOrDefault();
        if (paragraph == null)
        {
            paragraph = new Paragraph();
            cell.AppendChild(paragraph);
        }
        else
        {
            // Clear existing content but keep the paragraph
            paragraph.RemoveAllChildren();
        }
        
        // Create a new run with the content
        var run = new Run(new Text(newContent));
        paragraph.AppendChild(run);
    }
}