using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ForgeDoc.Structs;

namespace ForgeDoc.Processors;

public class WordTemplate
{
    private readonly string _templatePath;
    private readonly WordTemplateData _data;

    public WordTemplate(string templatePath, WordTemplateData data)
    {
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException("Template file not found", templatePath);
        }
        
        _templatePath = templatePath;
        _data = data;
    }

    public byte[] GetFile()
    {
        try
        {
            using (MemoryStream mem = new MemoryStream())
            {
                // Copy template to memory stream
                using (FileStream fileStream = new FileStream(_templatePath, FileMode.Open, FileAccess.Read))
                {
                    fileStream.CopyTo(mem);
                }
                mem.Position = 0; // Reset position after copying

                using (WordprocessingDocument doc = WordprocessingDocument.Open(mem, true))
                {
                    // Replace placeholders in main document
                    ReplacePlaceholders(doc.MainDocumentPart);

                    // Replace placeholders in headers
                    if (doc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        {
                            ReplacePlaceholdersInPart(headerPart);
                        }
                    }

                    // Replace placeholders in footers
                    if (doc.MainDocumentPart.FooterParts != null)
                    {
                        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        {
                            ReplacePlaceholdersInPart(footerPart);
                        }
                    }

                    // Process tables if any exist
                    ProcessTables(doc.MainDocumentPart);

                    doc.MainDocumentPart.Document.Save();
                }

                return mem.ToArray();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error processing Word template: {ex.Message}", ex);
        }
    }

    private void ReplacePlaceholders(MainDocumentPart mainPart)
    {
        if (mainPart?.Document?.Body == null) return;
        ReplacePlaceholdersInPart(mainPart);
    }

    private void ReplacePlaceholdersInPart(OpenXmlPart part)
    {
        if (part?.RootElement == null) return;

        // First, handle paragraphs that might contain split placeholders
        foreach (var paragraph in part.RootElement.Descendants<Paragraph>())
        {
            // Get all runs and their text content
            var runs = paragraph.Elements<Run>().ToList();
            if (!runs.Any()) continue;

            // Combine all text in the paragraph to check for placeholders
            string combinedText = string.Join("", runs.Select(r => 
                string.Join("", r.Elements<Text>().Select(t => t.Text))));

            bool containsPlaceholder = false;
            string modifiedText = combinedText;

            // Check if the combined text contains any placeholders
            foreach (var placeholder in _data.Placeholders)
            {
                string key = $"{{{{{placeholder.Key}}}}}";
                if (modifiedText.Contains(key))
                {
                    modifiedText = modifiedText.Replace(key, placeholder.Value ?? string.Empty);
                    containsPlaceholder = true;
                }
            }

            // If we found and replaced any placeholders, update the paragraph
            if (containsPlaceholder)
            {
                // Clear existing runs
                paragraph.RemoveAllChildren<Run>();

                // Add a new run with the modified text
                paragraph.AppendChild(new Run(new Text(modifiedText)));
            }
            else
            {
                // If no placeholders were found in the combined text,
                // still process individual text elements for partial matches
                foreach (var run in runs)
                {
                    foreach (var text in run.Elements<Text>())
                    {
                        string originalText = text.Text;
                        string textModified = originalText;

                        foreach (var placeholder in _data.Placeholders)
                        {
                            string key = $"{{{{{placeholder.Key}}}}}";
                            if (textModified.Contains(key))
                            {
                                textModified = textModified.Replace(key, placeholder.Value ?? string.Empty);
                            }
                        }

                        if (originalText != textModified)
                        {
                            text.Text = textModified;
                        }
                    }
                }
            }
        }
    }

    private void ProcessTables(MainDocumentPart mainPart)
    {
        if (mainPart?.Document?.Body == null) return;
        
        var tables = mainPart.Document.Body.Elements<Table>().ToList();
        
        foreach (var table in tables)
        {
            // Get the first row which typically contains headers
            var firstRow = table.Elements<TableRow>().FirstOrDefault();
            if (firstRow == null) continue;

            // Check if this table has any of our table placeholders
            var headerCells = firstRow.Elements<TableCell>().ToList();
            string tableIdentifier = null;

            foreach (var cell in headerCells)
            {
                var text = cell.Descendants<Text>().FirstOrDefault()?.Text;
                if (text != null && text.StartsWith("{{") && text.EndsWith("}}"))
                {
                    tableIdentifier = text.Trim('{', '}');
                    break;
                }
            }

            if (tableIdentifier != null && _data.Tables.ContainsKey(tableIdentifier))
            {
                var tableData = _data.Tables[tableIdentifier];
                
                // Remove the header row if it only contained our identifier
                if (headerCells.Count == 1)
                {
                    firstRow.Remove();
                }

                // Create a single row for the dictionary data
                var newRow = new TableRow();
                
                // Add a cell for the key
                var keyCell = new TableCell(
                    new Paragraph(
                        new Run(
                            new Text("Key")
                        )
                    )
                );
                newRow.Append(keyCell);
                
                // Add a cell for the value
                var valueCell = new TableCell(
                    new Paragraph(
                        new Run(
                            new Text("Value")
                        )
                    )
                );
                newRow.Append(valueCell);
                
                table.Append(newRow);

                // Add data rows
                foreach (var kvp in tableData)
                {
                    var dataRow = new TableRow();
                    
                    // Add key cell
                    var keyDataCell = new TableCell(
                        new Paragraph(
                            new Run(
                                new Text(kvp.Key ?? string.Empty)
                            )
                        )
                    );
                    dataRow.Append(keyDataCell);
                    
                    // Add value cell
                    var valueDataCell = new TableCell(
                        new Paragraph(
                            new Run(
                                new Text(kvp.Value ?? string.Empty)
                            )
                        )
                    );
                    dataRow.Append(valueDataCell);
                    
                    table.Append(dataRow);
                }
            }
        }
    }
}