using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
                    
                    // Process tables in main document
                    ProcessTablePlaceholders(doc.MainDocumentPart);

                    // Replace placeholders in headers
                    if (doc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        {
                            ReplacePlaceholdersInPart(headerPart);
                            ProcessTablePlaceholders(headerPart);
                        }
                    }

                    // Replace placeholders in footers
                    if (doc.MainDocumentPart.FooterParts != null)
                    {
                        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        {
                            ReplacePlaceholdersInPart(footerPart);
                            ProcessTablePlaceholders(footerPart);
                        }
                    }
                    
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
    
    private void ProcessTablePlaceholders(OpenXmlPart part)
    {
        if (part?.RootElement == null) return;
        
        // First, check for tables that already exist in the document
        // This should run even if no table data is provided
        ProcessExistingTables(part);
        
        // Then process the standard table placeholders
        // Only if table data is provided
        if (_data.Tables != null && _data.Tables.Any())
        {
            ProcessStandardTablePlaceholders(part);
        }
    }
    
    private void ProcessExistingTables(OpenXmlPart part)
    {
        // Find all tables in the document
        var tables = part.RootElement.Descendants<Table>().ToList();
        
        // Skip if no tables found or no table data provided
        if (!tables.Any() || _data.Tables == null || !_data.Tables.Any()) return;
        
        foreach (var table in tables)
        {
            // First, identify if this table contains table placeholders
            string tableName = null;
            List<Dictionary<string, string>> tableData = null;
            
            // Process each cell in the table to find table placeholders
            foreach (var row in table.Elements<TableRow>())
            {
                foreach (var cell in row.Elements<TableCell>())
                {
                    // Process each paragraph in the cell
                    foreach (var paragraph in cell.Elements<Paragraph>())
                    {
                        string cellText = GetParagraphText(paragraph);
                        
                        // Check for table start tag in the cell
                        foreach (var data in _data.Tables)
                        {
                            string startTag = $"{{{{#docTable {data.Key}}}}}";
                            if (cellText.Contains(startTag))
                            {
                                tableName = data.Key;
                                tableData = data.Value;
                                break;
                            }
                        }
                        
                        if (tableName != null) break;
                    }
                    
                    if (tableName != null) break;
                }
                
                if (tableName != null) break;
            }
            
            // If we found a table placeholder, process the entire table
            if (tableName != null && tableData != null)
            {
                ProcessExistingTableWithPlaceholders(table, tableName, tableData);
            }
        }
    }
    
    private void ProcessExistingTableWithPlaceholders(Table table, string tableName, List<Dictionary<string, string>> tableData)
    {
        // Get the row that contains the placeholders (usually the second row, after the header)
        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count < 2) return; // Need at least a header row and a data row
        
        // Find the row with the table placeholder
        TableRow templateRow = null;
        int templateRowIndex = -1;
        
        for (int i = 0; i < rows.Count; i++)
        {
            var row = rows[i];
            bool containsPlaceholder = false;
            
            // Check each cell in the row for the table placeholder
            foreach (var cell in row.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    string cellText = GetParagraphText(paragraph);
                    if (cellText.Contains($"{{{{#docTable {tableName}}}}}"))
                    {
                        containsPlaceholder = true;
                        break;
                    }
                }
                
                if (containsPlaceholder) break;
            }
            
            if (containsPlaceholder)
            {
                templateRow = row;
                templateRowIndex = i;
                break;
            }
        }
        
        // If we didn't find a row with the placeholder, use the second row as a fallback
        if (templateRow == null && rows.Count >= 2)
        {
            templateRow = rows[1];
            templateRowIndex = 1;
        }
        
        if (templateRow == null) return; // No template row found
        
        // Store a clone of the original template row before any processing
        TableRow originalTemplateRow = (TableRow)templateRow.CloneNode(true);
        
        // Keep a reference to the last row added
        TableRow lastRowAdded = templateRow;
        
        // Process each data item
        for (int i = 0; i < tableData.Count; i++)
        {
            var dataItem = tableData[i];
            
            // For the first item, we'll use the existing template row
            // For subsequent items, we'll create a new row from the original template
            TableRow newRow;
            if (i == 0)
            {
                newRow = templateRow;
            }
            else
            {
                // Clone the original template row (with placeholders)
                newRow = (TableRow)originalTemplateRow.CloneNode(true);
                
                // Add the new row to the table after the last row we added
                table.InsertAfter(newRow, lastRowAdded);
                lastRowAdded = newRow;
            }
            
            // Replace placeholders in each cell with the current data item
            foreach (var cell in newRow.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    string cellText = GetParagraphText(paragraph);
                    string processedText = cellText;
                    bool replacementMade = false;
                    
                    // First, remove any table start tags
                    var startTagPattern = new Regex(@"\{\{#docTable\s+[^}]+\}\}");
                    if (startTagPattern.IsMatch(processedText))
                    {
                        processedText = startTagPattern.Replace(processedText, "");
                        replacementMade = true;
                    }
                    
                    // Then, remove any table end tags
                    if (processedText.Contains("{{/docTable}}"))
                    {
                        processedText = processedText.Replace("{{/docTable}}", "");
                        replacementMade = true;
                    }
                    
                    // Check for standard placeholders {{Name}}
                    var standardPlaceholderPattern = new Regex(@"\{\{([^}]+)\}\}");
                    var matches = standardPlaceholderPattern.Matches(processedText);
                    
                    foreach (Match match in matches)
                    {
                        string placeholder = match.Groups[1].Value;
                        
                        // Skip if this is a table start tag
                        if (placeholder.StartsWith("#docTable")) continue;
                        
                        // Remove the end tag if found
                        if (placeholder == "/docTable" || placeholder.Trim() == "/docTable")
                        {
                            processedText = processedText.Replace($"{{{{{placeholder}}}}}", "");
                            replacementMade = true;
                            continue;
                        }
                        
                        // Check if this is an item placeholder (item.property)
                        if (placeholder.StartsWith("item."))
                        {
                            string itemProperty = placeholder.Substring(5); // Remove "item." prefix
                            if (dataItem.ContainsKey(itemProperty))
                            {
                                processedText = processedText.Replace($"{{{{{placeholder}}}}}", dataItem[itemProperty]);
                                replacementMade = true;
                            }
                        }
                        // Check if this is a direct property name
                        else if (dataItem.ContainsKey(placeholder))
                        {
                            processedText = processedText.Replace($"{{{{{placeholder}}}}}", dataItem[placeholder]);
                            replacementMade = true;
                        }
                    }
                    
                    // If we made any replacements, update the paragraph text
                    if (replacementMade)
                    {
                        // Final cleanup of any remaining table tags
                        processedText = Regex.Replace(processedText, @"\{\{#docTable\s+[^}]+\}\}", "");
                        processedText = processedText.Replace("{{/docTable}}", "");
                        processedText = processedText.Trim();
                        
                        ReplaceParagraphText(paragraph, processedText);
                    }
                }
            }
        }
    }
    
    private void DuplicateRowsForTableData(TableRow originalRow, List<Dictionary<string, string>> tableData)
    {
        if (tableData == null || tableData.Count <= 1) return;
        
        // Find the parent table
        Table parentTable = originalRow.Ancestors<Table>().FirstOrDefault();
        if (parentTable == null) return;
        
        // Create a copy of the original row for each additional data item
        for (int i = 1; i < tableData.Count; i++)
        {
            // Clone the original row
            TableRow newRow = (TableRow)originalRow.CloneNode(true);
            
            // Replace placeholders in the new row
            foreach (var cell in newRow.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    string paragraphText = GetParagraphText(paragraph);
                    string processedText = paragraphText;
                    bool replacementMade = false;
                    
                    // First, remove any table start tags
                    var startTagPattern = new Regex(@"\{\{#docTable\s+[^}]+\}\}");
                    if (startTagPattern.IsMatch(processedText))
                    {
                        processedText = startTagPattern.Replace(processedText, "");
                        replacementMade = true;
                    }
                    
                    // Then, remove any table end tags
                    if (processedText.Contains("{{/docTable}}"))
                    {
                        processedText = processedText.Replace("{{/docTable}}", "");
                        replacementMade = true;
                    }
                    
                    // Check for standard placeholders {{Name}}
                    var standardPlaceholderPattern = new Regex(@"\{\{([^}]+)\}\}");
                    var matches = standardPlaceholderPattern.Matches(processedText);
                    
                    foreach (Match match in matches)
                    {
                        string placeholder = match.Groups[1].Value;
                        
                        // Skip if this is a table start tag
                        if (placeholder.StartsWith("#docTable")) continue;
                        
                        // Remove the end tag if found
                        if (placeholder == "/docTable" || placeholder.Trim() == "/docTable")
                        {
                            processedText = processedText.Replace($"{{{{{placeholder}}}}}", "");
                            replacementMade = true;
                            continue;
                        }
                        
                        // Check if this is an item placeholder (item.property)
                        if (tableData[i].ContainsKey(placeholder))
                        {
                            // Replace with the current row's value
                            processedText = processedText.Replace($"{{{{{placeholder}}}}}", tableData[i][placeholder]);
                            replacementMade = true;
                        }
                    }
                    
                    // If we made any replacements, update the paragraph text
                    if (replacementMade)
                    {
                        // Final cleanup of any remaining table tags
                        processedText = Regex.Replace(processedText, @"\{\{#docTable\s+[^}]+\}\}", "");
                        processedText = processedText.Replace("{{/docTable}}", "");
                        processedText = processedText.Trim();
                        
                        ReplaceParagraphText(paragraph, processedText);
                    }
                }
            }
            
            // Add the new row to the table
            parentTable.InsertAfter(newRow, originalRow);
            
            // Update the original row reference to the newly added row
            // This ensures that each new row is added after the previous one
            originalRow = newRow;
        }
    }
    
    private void ReplaceCellContent(TableCell cell, List<Dictionary<string, string>> tableData)
    {
        // Clear the existing content
        cell.RemoveAllChildren();
        
        // Add a paragraph for each row in the table data
        foreach (var rowData in tableData)
        {
            Paragraph paragraph = new Paragraph();
            Run run = new Run();
            
            // Create text with all the data
            string text = string.Join(", ", rowData.Select(kv => $"{kv.Key}: {kv.Value}"));
            run.AppendChild(new Text(text));
            
            paragraph.AppendChild(run);
            cell.AppendChild(paragraph);
        }
    }
    
    private void ReplaceParagraphText(Paragraph paragraph, string newText)
    {
        // Clear existing runs
        paragraph.RemoveAllChildren();
        
        // Add a new run with the new text
        Run run = new Run();
        run.AppendChild(new Text(newText));
        paragraph.AppendChild(run);
    }
    
    private void ProcessStandardTablePlaceholders(OpenXmlPart part)
    {
        // Get all paragraphs in the document
        var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();
        
        foreach (var tableData in _data.Tables)
        {
            string tableName = tableData.Key;
            List<Dictionary<string, string>> data = tableData.Value;
            
            // Skip if no data
            if (data == null || !data.Any()) continue;
            
            string startTag = $"{{{{#docTable {tableName}}}}}";
            string endTag = "{{/docTable}}";
            
            int startIndex = -1;
            int endIndex = -1;
            
            // Find the start and end tags in the paragraphs
            for (int i = 0; i < paragraphs.Count; i++)
            {
                var paragraph = paragraphs[i];
                string paragraphText = GetParagraphText(paragraph);
                
                if (paragraphText.Contains(startTag))
                {
                    startIndex = i;
                }
                
                if (paragraphText.Contains(endTag) && startIndex != -1 && i >= startIndex)
                {
                    endIndex = i;
                    break;
                }
            }
            
            // If we found both start and end tags
            if (startIndex != -1 && endIndex != -1)
            {
                // Extract the template content between the tags
                var templateContent = new StringBuilder();
                for (int i = startIndex; i <= endIndex; i++)
                {
                    templateContent.AppendLine(GetParagraphText(paragraphs[i]));
                }
                
                // Create the table
                Table table = CreateTable(templateContent.ToString(), tableData.Value);
                
                // Insert the table after the start paragraph
                if (table != null)
                {
                    paragraphs[startIndex].Parent.InsertAfter(table, paragraphs[startIndex]);
                }
                
                // Remove the paragraphs that contained the table template
                for (int i = endIndex; i >= startIndex; i--)
                {
                    paragraphs[i].Remove();
                }
            }
            // If we only found the start tag but not the end tag
            else if (startIndex != -1)
            {
                // Check if the start and end tags are in the same paragraph
                string paragraphText = GetParagraphText(paragraphs[startIndex]);
                int startTagIndex = paragraphText.IndexOf(startTag);
                int endTagIndex = paragraphText.IndexOf(endTag);
                
                if (startTagIndex != -1 && endTagIndex != -1 && endTagIndex > startTagIndex)
                {
                    // Extract the template content between the tags
                    string templateContent = paragraphText.Substring(
                        startTagIndex + startTag.Length,
                        endTagIndex - startTagIndex - startTag.Length);
                    
                    // Create the table
                    Table table = CreateTable(templateContent, tableData.Value);
                    
                    // Insert the table after the paragraph
                    paragraphs[startIndex].Parent.InsertAfter(table, paragraphs[startIndex]);
                    
                    // Remove the original paragraph
                    paragraphs[startIndex].Remove();
                }
            }
        }
    }
    
    private string GetParagraphText(Paragraph paragraph)
    {
        // Get all runs and their text content
        var runs = paragraph.Elements<Run>().ToList();
        if (!runs.Any()) return string.Empty;
        
        // Combine all text in the paragraph
        return string.Join("", runs.Select(r => 
            string.Join("", r.Elements<Text>().Select(t => t.Text))));
    }
    
    private Table CreateTable(string templateContent, List<Dictionary<string, string>> tableData)
    {
        if (string.IsNullOrWhiteSpace(templateContent) || tableData == null || !tableData.Any())
            return null;
        
        // Remove the table start and end tags if present
        templateContent = Regex.Replace(templateContent, @"\{\{#docTable\s+[^}]+\}\}", "");
        templateContent = templateContent.Replace("{{/docTable}}", "");
        
        // Extract all unique placeholders in format {{item.xxx}}
        var placeholderPattern = new Regex(@"\{\{item\.([^}]+)\}\}");
        var matches = placeholderPattern.Matches(templateContent);
        
        // Get unique column names
        var columnNames = new List<string>();
        foreach (Match match in matches)
        {
            string columnName = match.Groups[1].Value;
            if (!columnNames.Contains(columnName))
            {
                columnNames.Add(columnName);
            }
        }
        
        // If no columns found, try to use the keys from the first data item
        if (!columnNames.Any() && tableData.Any() && tableData[0].Any())
        {
            columnNames.AddRange(tableData[0].Keys);
        }
        
        // Create the table
        Table table = new Table();
        
        // Set table properties
        TableProperties tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 },
                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 1 }
            ),
            new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct }
        );
        table.AppendChild(tableProperties);
        
        // Add header row
        TableRow headerRow = new TableRow();
        foreach (var columnName in columnNames)
        {
            headerRow.Append(CreateTableCell(columnName, true));
        }
        table.Append(headerRow);
        
        // Add data rows
        foreach (var dataItem in tableData)
        {
            TableRow dataRow = new TableRow();
            
            foreach (var columnName in columnNames)
            {
                string cellValue = dataItem.ContainsKey(columnName) ? dataItem[columnName] : "";
                dataRow.Append(CreateTableCell(cellValue, false));
            }
            
            table.Append(dataRow);
        }
        
        return table;
    }
    
    private TableCell CreateTableCell(string text, bool isHeader)
    {
        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Auto }
            ),
            new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = JustificationValues.Center }
                ),
                new Run(
                    isHeader ? new RunProperties(new Bold()) : null,
                    new Text(text)
                )
            )
        );
    }
    
    private Drawing AddImageToAnyWhere(WordprocessingDocument document, string relationshipId, int width, int height)
    {
        var nextId = GetMaxDocPropertyId(document) + 1;

        Size size = new Size(width, height); //Size(260, 80);

        long widthInEmus = size.Width * 9525;
        long heightInEmus = size.Height * 9525;

        // Define the reference of the image.
        var element =
            new Drawing(
                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent() { Cx = widthInEmus, Cy = heightInEmus },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties()
                    {
                        Id = (UInt32Value)nextId,
                        Name = string.Format("Picture {0}", nextId)
                    },
                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties(
                        new DocumentFormat.OpenXml.Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                    new DocumentFormat.OpenXml.Drawing.Graphic(
                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties()
                                    {
                                        Id = (UInt32Value)0U,
                                        Name = "New Bitmap Image.jpg"
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()),
                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                    new DocumentFormat.OpenXml.Drawing.Blip(
                                        new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                                            new DocumentFormat.OpenXml.Drawing.BlipExtension()
                                            {
                                                Uri =
                                                    "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            })
                                        )
                                    {
                                        Embed = relationshipId,
                                        CompressionState =
                                            DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                        new DocumentFormat.OpenXml.Drawing.FillRectangle())),
                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                        new DocumentFormat.OpenXml.Drawing.Offset() { X = 0L, Y = 0L },
                                        new DocumentFormat.OpenXml.Drawing.Extents() { Cx = widthInEmus, Cy = heightInEmus }),
                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                        )
                                    { Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle }))
                            )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                    )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = "50D07946"
                });

        // Append the reference to body, the element should be in a Run.
        // var newParagraph = new Paragraph(new Run(element));
        //ParagraphHelper.AddToBody(newParagraph);
        //ParagraphHelper.AddToBody(newParagraph);

        //return newParagraph;
        return element;
    }
    
    private uint GetMaxDocPropertyId(WordprocessingDocument document)
    {
        return document
            .MainDocumentPart
            .RootElement
            .Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .Max(x => (uint?)x.Id) ?? 0;
    }
}