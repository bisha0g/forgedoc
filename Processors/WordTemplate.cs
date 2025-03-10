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
        // Find tables with docTable markers
        var tables = body.Descendants<Table>().ToList();
        foreach (var table in tables)
        {
            // Check if any cell contains our docTable marker
            bool markerFound = false;
            foreach (var cell in table.Descendants<TableCell>())
            {
                var paragraphs = cell.Descendants<Paragraph>().ToList();
                foreach (var paragraph in paragraphs)
                {
                    var text = paragraph.InnerText;
                    if (text.Contains($"{{{{#docTable {tableVariable}}}}}"))
                    {
                        // Process the table with the marker
                        ProcessTableData(table, tableVariable, tableData);
                        markerFound = true;
                        break;
                    }
                }
                if (markerFound) break;
            }
            if (markerFound) return;
        }
        
        // If we get here, we didn't find the marker inside a table cell
        // Now look for the marker in paragraphs before tables
        var allParagraphs = body.Descendants<Paragraph>().ToList();
        for (int i = 0; i < allParagraphs.Count; i++)
        {
            var paragraphText = allParagraphs[i].InnerText;
            if (paragraphText.Contains($"{{{{#docTable {tableVariable}}}}}"))
            {
                // Find the associated table by looking at subsequent elements
                var currentElement = allParagraphs[i];
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
                    allParagraphs[i].Remove();
                    
                    // Process the table
                    ProcessTableData(table, tableVariable, tableData);
                    return; // Table found and processed
                }
            }
        }
    }
    
    private void ProcessTableData(Table table, string tableVariable, List<Dictionary<string, string>> tableData)
    {
        // Get all rows in the table
        var rows = table.Elements<TableRow>().ToList();
        
        // We need at least one row as a template
        if (rows.Count == 0)
            return;
            
        // Find the template row and end marker row
        TableRow? templateRow = null;
        int templateRowIndex = -1;
        int endMarkerRowIndex = -1;
        
        for (int i = 0; i < rows.Count; i++)
        {
            var rowText = rows[i].InnerText;
            
            // Check for template row with docTable marker
            if (rowText.Contains($"{{{{#docTable {tableVariable}}}}}"))
            {
                templateRow = rows[i];
                templateRowIndex = i;
            }
            // Check for end marker
            else if (rowText.Contains($"{{{{/docTable}}}}"))
            {
                endMarkerRowIndex = i;
                // If the end marker is in the same row as template, we'll use this row as template
                if (templateRowIndex == -1)
                {
                    templateRow = rows[i];
                    templateRowIndex = i;
                }
            }
        }
        
        // If we couldn't find the template row, use the first row
        if (templateRow == null)
        {
            templateRow = rows[0];
            templateRowIndex = 0;
        }
        
        // Clone the template row to preserve styling
        var clonedTemplateRow = (TableRow)templateRow.CloneNode(true);
        
        // Clean up the template row by removing all docTable markers
        RemoveAllTableMarkers(clonedTemplateRow, tableVariable);
        
        // Store table properties before modifying rows
        var tableProperties = table.GetFirstChild<TableProperties>()?.CloneNode(true);
        
        // Determine which rows to keep (those outside the docTable section)
        var rowsToKeep = new List<TableRow>();
        for (int i = 0; i < rows.Count; i++)
        {
            if (i < templateRowIndex || (endMarkerRowIndex != -1 && i > endMarkerRowIndex))
            {
                var rowToKeep = (TableRow)rows[i].CloneNode(true);
                // Also clean up any markers from rows we're keeping
                RemoveAllTableMarkers(rowToKeep, tableVariable);
                rowsToKeep.Add(rowToKeep);
            }
        }
        
        // Remove all existing rows but preserve table properties
        table.RemoveAllChildren();
        
        // Re-add table properties if they existed
        if (tableProperties != null)
        {
            table.AppendChild(tableProperties);
        }
        
        // Add back rows that were before the template
        for (int i = 0; i < templateRowIndex; i++)
        {
            table.AppendChild(rowsToKeep[i]);
        }
        
        // Create rows for each data item using the template
        foreach (var rowData in tableData)
        {
            // Clone the template row to preserve styling
            var newRow = (TableRow)clonedTemplateRow.CloneNode(true);
            
            // Replace placeholders in each cell
            ReplaceDataPlaceholders(newRow, rowData);
            
            table.AppendChild(newRow);
        }
        
        // Add back rows that were after the end marker
        if (endMarkerRowIndex != -1)
        {
            for (int i = endMarkerRowIndex + 1; i < rows.Count; i++)
            {
                int keepIndex = i - (endMarkerRowIndex - templateRowIndex);
                if (keepIndex < rowsToKeep.Count)
                {
                    table.AppendChild(rowsToKeep[keepIndex]);
                }
            }
        }
        
        // Final scan through the entire table to remove any remaining markers
        foreach (var row in table.Elements<TableRow>())
        {
            RemoveAllTableMarkers(row, tableVariable);
        }
    }
    
    // Helper method to remove all table markers from a row
    private void RemoveAllTableMarkers(TableRow row, string tableVariable)
    {
        // Remove start marker
        CleanupTableRow(row, $"{{{{#docTable {tableVariable}}}}}", "");
        // Remove end marker
        CleanupTableRow(row, "{{/docTable}}", "");
    }
    
    // Helper method to replace data placeholders in a table row
    private void ReplaceDataPlaceholders(TableRow row, Dictionary<string, string> data)
    {
        foreach (var cell in row.Descendants<TableCell>())
        {
            foreach (var paragraph in cell.Descendants<Paragraph>())
            {
                foreach (var run in paragraph.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        string originalText = text.Text;
                        
                        // Skip processing if the text is empty
                        if (string.IsNullOrEmpty(originalText))
                            continue;
                            
                        string newText = originalText;
                        
                        // First, remove any docTable markers that might be in this text element
                        if (newText.Contains("{{#docTable"))
                        {
                            int startIndex = newText.IndexOf("{{#docTable");
                            int endIndex = newText.IndexOf("}}", startIndex);
                            if (endIndex > startIndex)
                            {
                                newText = newText.Remove(startIndex, endIndex - startIndex + 2);
                            }
                        }
                        
                        if (newText.Contains("{{/docTable}}"))
                        {
                            newText = newText.Replace("{{/docTable}}", "");
                        }
                        
                        // Replace each data placeholder with its value
                        foreach (var kvp in data)
                        {
                            string placeholder = $"{{{{{kvp.Key}}}}}";
                            
                            // Use a more precise replacement approach
                            int startPos = 0;
                            while (true)
                            {
                                int pos = newText.IndexOf(placeholder, startPos);
                                if (pos < 0) break;
                                
                                // Check if this is a standalone placeholder and not part of another variable
                                bool isStandalone = true;
                                
                                // Check if it's preceded by a dot (which would indicate it's part of a property access)
                                if (pos > 0 && newText[pos - 1] == '.')
                                {
                                    isStandalone = false;
                                }
                                
                                if (isStandalone)
                                {
                                    // Replace this occurrence
                                    newText = newText.Substring(0, pos) + kvp.Value + newText.Substring(pos + placeholder.Length);
                                    startPos = pos + kvp.Value.Length;
                                }
                                else
                                {
                                    // Skip this occurrence
                                    startPos = pos + placeholder.Length;
                                }
                            }
                        }
                        
                        // Only update if changed
                        if (newText != originalText)
                        {
                            text.Text = newText;
                        }
                    }
                }
            }
        }
    }
    
    private void CleanupTableRow(TableRow row, string textToReplace, string replacement)
    {
        foreach (var cell in row.Descendants<TableCell>())
        {
            foreach (var paragraph in cell.Descendants<Paragraph>())
            {
                foreach (var run in paragraph.Descendants<Run>())
                {
                    foreach (var text in run.Descendants<Text>())
                    {
                        if (text.Text.Contains(textToReplace))
                        {
                            text.Text = text.Text.Replace(textToReplace, replacement);
                        }
                    }
                }
            }
        }
    }
}