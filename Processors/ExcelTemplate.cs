using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;
using System.Drawing.Imaging;
using ForgeDoc.Structs;
using A = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Color = DocumentFormat.OpenXml.Spreadsheet.Color;

namespace ForgeDoc.Processors;

public class ExcelTemplate
{
    private readonly string _templatePath;
    private readonly ExcelTemplateData _data;

    public ExcelTemplate(string templatePath, ExcelTemplateData data)
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

                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(mem, true))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    
                    // Process each worksheet
                    foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
                    {
                        // Replace placeholders in cells
                        ReplacePlaceholders(worksheetPart);
                        
                        // Process tables in worksheet
                        ProcessTablePlaceholders(worksheetPart);
                        
                        // Process image placeholders in worksheet
                        ProcessImagePlaceholders(worksheetPart);
                        
                        // Process for loop placeholders
                        ProcessForLoopPlaceholders(worksheetPart);
                    }
                    
                    workbookPart.Workbook.Save();
                }

                return mem.ToArray();
            }
        }
        catch (Exception ex)
        {
            throw new Exception($"Error processing Excel template: {ex.Message}", ex);
        }
    }

    private void ReplacePlaceholders(WorksheetPart worksheetPart)
    {
        if (worksheetPart?.Worksheet == null) return;
        
        // Get all cells in the worksheet
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        
        // Get shared string table
        SharedStringTablePart sharedStringTablePart = worksheetPart.GetParentParts()
            .OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart;
        
        if (sharedStringTablePart == null) return;
        
        // Process each cell in the worksheet
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                // Only process cells with shared string values
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int sharedStringIndex;
                    if (int.TryParse(cell.InnerText, out sharedStringIndex))
                    {
                        var sharedString = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                        string cellText = sharedString.InnerText;
                        bool modified = false;
                        bool hasRichText = false;
                        
                        // Replace placeholders with rich text formatting
                        foreach (var placeholder in _data.RichTextPlaceholders)
                        {
                            string key = $"{{{{{placeholder.Key}}}}}";
                            if (cellText.Contains(key))
                            {
                                // For Excel, we need to convert HTML-like tags to OpenXML formatting
                                string richText = placeholder.Value;
                                cellText = cellText.Replace(key, richText);
                                modified = true;
                                hasRichText = true;
                            }
                        }
                        
                        // Replace regular placeholders
                        foreach (var placeholder in _data.Placeholders)
                        {
                            string key = $"{{{{{placeholder.Key}}}}}";
                            if (cellText.Contains(key))
                            {
                                cellText = cellText.Replace(key, placeholder.Value ?? string.Empty);
                                modified = true;
                            }
                        }
                        
                        // Replace special characters
                        foreach (var specialChar in _data.SpecialCharacters)
                        {
                            string key = $"{{{{{specialChar.Key}}}}}";
                            if (cellText.Contains(key))
                            {
                                cellText = cellText.Replace(key, specialChar.Value.Character);
                                modified = true;
                                // Note: In Excel, we can't easily change fonts for individual characters
                                // so we just replace with the character
                            }
                        }
                        
                        // If the cell was modified, update the shared string
                        if (modified)
                        {
                            SharedStringItem newSharedString;
                            
                            if (hasRichText)
                            {
                                // Create a rich text shared string
                                newSharedString = CreateRichTextSharedString(cellText);
                            }
                            else
                            {
                                // Create a simple text shared string
                                newSharedString = new SharedStringItem(new Text(cellText));
                            }
                            
                            // Add to shared string table and get the index
                            sharedStringTablePart.SharedStringTable.AppendChild(newSharedString);
                            int newIndex = sharedStringTablePart.SharedStringTable.Count() - 1;
                            
                            // Update the cell to point to the new shared string
                            cell.CellValue = new CellValue(newIndex.ToString());
                        }
                    }
                }
                else if (cell.CellValue != null)
                {
                    // Handle direct cell values (not shared strings)
                    string cellText = cell.CellValue.Text;
                    bool modified = false;
                    bool hasRichText = false;
                    
                    // Replace placeholders with rich text formatting
                    foreach (var placeholder in _data.RichTextPlaceholders)
                    {
                        string key = $"{{{{{placeholder.Key}}}}}";
                        if (cellText.Contains(key))
                        {
                            // For direct cell values, we need to convert to shared string to support rich text
                            string richText = placeholder.Value;
                            cellText = cellText.Replace(key, richText);
                            modified = true;
                            hasRichText = true;
                        }
                    }
                    
                    // Replace regular placeholders
                    foreach (var placeholder in _data.Placeholders)
                    {
                        string key = $"{{{{{placeholder.Key}}}}}";
                        if (cellText.Contains(key))
                        {
                            cellText = cellText.Replace(key, placeholder.Value ?? string.Empty);
                            modified = true;
                        }
                    }
                    
                    // Replace special characters
                    foreach (var specialChar in _data.SpecialCharacters)
                    {
                        string key = $"{{{{{specialChar.Key}}}}}";
                        if (cellText.Contains(key))
                        {
                            cellText = cellText.Replace(key, specialChar.Value.Character);
                            modified = true;
                        }
                    }
                    
                    // If the cell was modified, update the cell value
                    if (modified)
                    {
                        if (hasRichText)
                        {
                            // Convert to shared string for rich text
                            SharedStringItem newSharedString = CreateRichTextSharedString(cellText);
                            
                            // Add to shared string table and get the index
                            sharedStringTablePart.SharedStringTable.AppendChild(newSharedString);
                            int newIndex = sharedStringTablePart.SharedStringTable.Count() - 1;
                            
                            // Update cell to use shared string
                            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                            cell.CellValue = new CellValue(newIndex.ToString());
                        }
                        else
                        {
                            // Update regular cell value
                            cell.CellValue = new CellValue(cellText);
                        }
                    }
                }
            }
        }
        
        // Save the shared string table
        if (sharedStringTablePart != null)
        {
            sharedStringTablePart.SharedStringTable.Save();
        }
    }
    
    private SharedStringItem CreateRichTextSharedString(string text)
    {
        var sharedStringItem = new SharedStringItem();
        
        // Process HTML-like formatting tags
        // We'll handle <b>, <i>, <u>, and <span style="color:#RRGGBB">
        
        // Parse the text and create runs with appropriate formatting
        int currentPosition = 0;
        
        while (currentPosition < text.Length)
        {
            int tagStart = text.IndexOf('<', currentPosition);
            
            if (tagStart == -1)
            {
                // No more tags, add the remaining text
                if (currentPosition < text.Length)
                {
                    sharedStringItem.AppendChild(new Run(new Text(text.Substring(currentPosition))));
                }
                break;
            }
            
            // Add text before the tag
            if (tagStart > currentPosition)
            {
                sharedStringItem.AppendChild(new Run(new Text(text.Substring(currentPosition, tagStart - currentPosition))));
            }
            
            int tagEnd = text.IndexOf('>', tagStart);
            if (tagEnd == -1)
            {
                // Malformed tag, treat as text
                sharedStringItem.AppendChild(new Run(new Text(text.Substring(tagStart))));
                break;
            }
            
            string tagName = text.Substring(tagStart + 1, tagEnd - tagStart - 1).Split(' ')[0].ToLower();
            bool isClosingTag = tagName.StartsWith("/");
            
            if (isClosingTag)
            {
                tagName = tagName.Substring(1); // Remove the '/'
            }
            
            // Find the content between opening and closing tags
            if (!isClosingTag)
            {
                int closingTagStart = text.IndexOf("</" + tagName + ">", tagEnd + 1);
                if (closingTagStart == -1)
                {
                    // No closing tag, treat as text
                    sharedStringItem.AppendChild(new Run(new Text(text.Substring(tagStart, tagEnd - tagStart + 1))));
                    currentPosition = tagEnd + 1;
                    continue;
                }
                
                string content = text.Substring(tagEnd + 1, closingTagStart - tagEnd - 1);
                
                // Create a run with appropriate formatting
                var run = new Run();
                var runProperties = new RunProperties();
                
                switch (tagName)
                {
                    case "b":
                        runProperties.AppendChild(new Bold());
                        break;
                    case "i":
                        runProperties.AppendChild(new Italic());
                        break;
                    case "u":
                        runProperties.AppendChild(new Underline());
                        break;
                    case "span":
                        // Check for color style
                        string tagAttributes = text.Substring(tagStart + 1 + "span".Length, tagEnd - tagStart - 1 - "span".Length);
                        var colorMatch = Regex.Match(tagAttributes, @"style\s*=\s*""color\s*:\s*#([0-9A-Fa-f]{6})""");
                        if (colorMatch.Success)
                        {
                            string colorCode = colorMatch.Groups[1].Value;
                            runProperties.AppendChild(new Color() { Rgb = new HexBinaryValue() { Value = colorCode } });
                        }
                        break;
                }
                
                run.AppendChild(runProperties);
                run.AppendChild(new Text(content));
                sharedStringItem.AppendChild(run);
                
                currentPosition = closingTagStart + tagName.Length + 3; // +3 for "</>"
            }
            else
            {
                // Skip closing tags as they are handled with opening tags
                currentPosition = tagEnd + 1;
            }
        }
        
        // If no rich text was added, add the plain text
        if (!sharedStringItem.HasChildren)
        {
            sharedStringItem.AppendChild(new Run(new Text(text)));
        }
        
        return sharedStringItem;
    }

    private void ProcessTablePlaceholders(WorksheetPart worksheetPart)
    {
        if (worksheetPart?.Worksheet == null || _data.Tables == null || !_data.Tables.Any()) return;
        
        // Get all cells in the worksheet
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        
        // Get shared string table
        SharedStringTablePart sharedStringTablePart = worksheetPart.GetParentParts()
            .OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart;
        
        if (sharedStringTablePart == null) return;
        
        // Find table placeholders in the worksheet
        Dictionary<string, (int RowIndex, int ColIndex, string TableName)> tablePlaceholders = new Dictionary<string, (int, int, string)>();
        
        // Process each cell to find table placeholders
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                // Get cell text
                string cellText = GetCellText(cell, sharedStringTablePart);
                if (string.IsNullOrEmpty(cellText)) continue;
                
                // Check for table placeholders
                var match = Regex.Match(cellText, @"\{\{#docTable\s+(\w+)\s*\}\}");
                if (match.Success)
                {
                    string tableName = match.Groups[1].Value;
                    
                    // Get cell reference
                    string cellReference = cell.CellReference.Value;
                    (int rowIndex, int colIndex) = GetCellPosition(cellReference);
                    
                    // Store the table placeholder
                    tablePlaceholders.Add(cellReference, (rowIndex, colIndex, tableName));
                }
            }
        }
        
        // Process each table placeholder
        foreach (var placeholder in tablePlaceholders)
        {
            string cellReference = placeholder.Key;
            (int rowIndex, int colIndex, string tableName) = placeholder.Value;
            
            // Check if the table exists in the data
            if (!_data.HasTable(tableName)) continue;
            
            // Get the table data
            var tableData = _data.Tables[tableName];
            if (tableData == null || !tableData.Any()) continue;
            
            // Get the header row to determine columns
            var headerRow = tableData.FirstOrDefault();
            if (headerRow == null) continue;
            
            // Create header row
            int currentRowIndex = rowIndex;
            InsertRow(sheetData, currentRowIndex, headerRow.Keys.ToList(), colIndex);
            currentRowIndex++;
            
            // Create data rows
            foreach (var dataRow in tableData)
            {
                InsertRow(sheetData, currentRowIndex, dataRow.Values.ToList(), colIndex);
                currentRowIndex++;
            }
            
            // Remove the placeholder cell
            var placeholderCell = GetCellByReference(sheetData, cellReference);
            if (placeholderCell != null)
            {
                placeholderCell.Remove();
            }
        }
        
        // Save the worksheet
        worksheetPart.Worksheet.Save();
    }
    
    private string GetCellText(Cell cell, SharedStringTablePart sharedStringTablePart)
    {
        if (cell == null || cell.CellValue == null) return string.Empty;
        
        // Handle shared strings
        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
        {
            int sharedStringIndex;
            if (int.TryParse(cell.CellValue.Text, out sharedStringIndex))
            {
                var sharedString = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);
                return sharedString.InnerText;
            }
        }
        
        // Handle direct values
        return cell.CellValue.Text;
    }
    
    private (int RowIndex, int ColIndex) GetCellPosition(string cellReference)
    {
        // Extract column letters and row number from cell reference (e.g., "A1")
        var match = Regex.Match(cellReference, @"([A-Z]+)(\d+)");
        if (!match.Success) return (0, 0);
        
        string colLetters = match.Groups[1].Value;
        int rowIndex = int.Parse(match.Groups[2].Value);
        
        // Convert column letters to index (A=1, B=2, etc.)
        int colIndex = 0;
        foreach (char c in colLetters)
        {
            colIndex = colIndex * 26 + (c - 'A' + 1);
        }
        
        return (rowIndex, colIndex);
    }
    
    private string GetCellReference(int rowIndex, int colIndex)
    {
        // Convert column index to letters (1=A, 2=B, etc.)
        string colLetters = string.Empty;
        while (colIndex > 0)
        {
            int remainder = (colIndex - 1) % 26;
            colLetters = (char)('A' + remainder) + colLetters;
            colIndex = (colIndex - 1) / 26;
        }
        
        return $"{colLetters}{rowIndex}";
    }
    
    private Cell GetCellByReference(SheetData sheetData, string cellReference)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference.Value == cellReference)
                {
                    return cell;
                }
            }
        }
        
        return null;
    }
    
    private void InsertRow(SheetData sheetData, int rowIndex, List<string> values, int startColIndex)
    {
        // Check if row already exists
        Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowIndex);
        if (row == null)
        {
            // Create new row
            row = new Row() { RowIndex = (uint)rowIndex };
            
            // Find the correct position to insert the row
            Row nextRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value > rowIndex);
            if (nextRow != null)
            {
                sheetData.InsertBefore(row, nextRow);
            }
            else
            {
                sheetData.AppendChild(row);
            }
        }
        
        // Add cells to the row
        for (int i = 0; i < values.Count; i++)
        {
            int colIndex = startColIndex + i;
            string cellReference = GetCellReference(rowIndex, colIndex);
            
            // Create cell
            Cell cell = new Cell() { CellReference = cellReference };
            
            // Set cell value
            cell.CellValue = new CellValue(values[i]);
            cell.DataType = CellValues.String;
            
            // Add cell to row
            row.AppendChild(cell);
        }
    }

    private void ProcessImagePlaceholders(WorksheetPart worksheetPart)
    {
        if (worksheetPart?.Worksheet == null || _data.Images == null || !_data.Images.Any()) return;
        
        // Get all cells in the worksheet
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        
        // Get shared string table
        SharedStringTablePart sharedStringTablePart = worksheetPart.GetParentParts()
            .OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart;
        
        if (sharedStringTablePart == null) return;
        
        // Find image placeholders in the worksheet
        Dictionary<string, (int RowIndex, int ColIndex, string ImageKey, int Width, int Height)> imagePlaceholders = new Dictionary<string, (int, int, string, int, int)>();
        
        // Process each cell to find image placeholders
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>())
            {
                // Get cell text
                string cellText = GetCellText(cell, sharedStringTablePart);
                if (string.IsNullOrEmpty(cellText)) continue;
                
                // Check for image placeholders in the format {% ImageKey %} or {% ImageKey:widthxheight %}
                var match = Regex.Match(cellText, @"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
                if (match.Success)
                {
                    string imageKey = match.Groups[1].Value.Trim();
                    
                    // Check if this image key exists in our data
                    if (!_data.HasImage(imageKey)) continue;
                    
                    // Get cell reference
                    string cellReference = cell.CellReference.Value;
                    (int rowIndex, int colIndex) = GetCellPosition(cellReference);
                    
                    // Get width and height if specified
                    int width = 400; // Default width
                    int height = 300; // Default height
                    
                    if (match.Groups.Count > 2 && !string.IsNullOrEmpty(match.Groups[2].Value))
                    {
                        int.TryParse(match.Groups[2].Value, out width);
                    }
                    
                    if (match.Groups.Count > 3 && !string.IsNullOrEmpty(match.Groups[3].Value))
                    {
                        int.TryParse(match.Groups[3].Value, out height);
                    }
                    
                    // Store the image placeholder
                    imagePlaceholders.Add(cellReference, (rowIndex, colIndex, imageKey, width, height));
                }
            }
        }
        
        // Process each image placeholder
        foreach (var placeholder in imagePlaceholders)
        {
            string cellReference = placeholder.Key;
            (int rowIndex, int colIndex, string imageKey, int width, int height) = placeholder.Value;
            
            // Get the image path
            string imagePath = _data.GetImage(imageKey);
            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath)) continue;
            
            // Insert the image
            InsertImage(worksheetPart, rowIndex, colIndex, imagePath, width, height);
            
            // Clear the placeholder text from the cell
            var placeholderCell = GetCellByReference(sheetData, cellReference);
            if (placeholderCell != null)
            {
                placeholderCell.CellValue = new CellValue(string.Empty);
                if (placeholderCell.DataType != null)
                {
                    placeholderCell.DataType = CellValues.String;
                }
            }
        }
    }
    
    private void InsertImage(WorksheetPart worksheetPart, int rowIndex, int colIndex, string imagePath, int maxWidth, int maxHeight)
    {
        try
        {
            // Get the DrawingsPart or create it if it doesn't exist
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null)
            {
                drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                var worksheetDrawing = new Xdr.WorksheetDrawing();
                worksheetDrawing.AddNamespaceDeclaration("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                worksheetDrawing.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
                drawingsPart.WorksheetDrawing = worksheetDrawing;
                
                // Add the drawing reference to the worksheet
                var drawing = new DocumentFormat.OpenXml.Spreadsheet.Drawing() { Id = worksheetPart.GetIdOfPart(drawingsPart) };
                worksheetPart.Worksheet.Append(drawing);
            }
            
            // Get image dimensions
            int imageWidthEmu;
            int imageHeightEmu;
            
            using (var img = System.Drawing.Image.FromFile(imagePath))
            {
                // Calculate new dimensions while maintaining aspect ratio
                double scale = 1.0;
                if (img.Width > maxWidth || img.Height > maxHeight)
                {
                    double widthScale = maxWidth / (double)img.Width;
                    double heightScale = maxHeight / (double)img.Height;
                    scale = Math.Min(widthScale, heightScale);
                }
                
                int newWidth = (int)(img.Width * scale);
                int newHeight = (int)(img.Height * scale);
                
                // Convert pixels to EMUs (English Metric Units)
                imageWidthEmu = newWidth * 9525;
                imageHeightEmu = newHeight * 9525;
            }
            
            // Add the image part
            ImagePart imagePart = drawingsPart.AddImagePart(GetImagePartTypeFromFormat(imagePath));
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            
            // Get the next available drawing ID
            uint drawingId = 1;
            if (drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>().Any())
            {
                drawingId = drawingsPart.WorksheetDrawing.Elements<Xdr.TwoCellAnchor>()
                    .Max(a => a.Elements<Xdr.Picture>().First().NonVisualPictureProperties.NonVisualDrawingProperties.Id.Value) + 1;
            }
            
            // Create the drawing
            var twoCellAnchor = new Xdr.TwoCellAnchor();
            
            // From position (top-left)
            var fromMarker = new Xdr.FromMarker();
            fromMarker.ColumnId = new Xdr.ColumnId() { Text = (colIndex - 1).ToString() };
            fromMarker.RowId = new Xdr.RowId() { Text = (rowIndex - 1).ToString() };
            fromMarker.ColumnOffset = new Xdr.ColumnOffset() { Text = "0" };
            fromMarker.RowOffset = new Xdr.RowOffset() { Text = "0" };
            twoCellAnchor.Append(fromMarker);
            
            // To position (bottom-right)
            var toMarker = new Xdr.ToMarker();
            // Calculate the column and row span based on the image size
            // This is a simplified calculation; you might need to adjust it based on your needs
            toMarker.ColumnId = new Xdr.ColumnId() { Text = colIndex.ToString() };
            toMarker.RowId = new Xdr.RowId() { Text = rowIndex.ToString() };
            toMarker.ColumnOffset = new Xdr.ColumnOffset() { Text = imageWidthEmu.ToString() };
            toMarker.RowOffset = new Xdr.RowOffset() { Text = imageHeightEmu.ToString() };
            twoCellAnchor.Append(toMarker);
            
            // Create the picture
            var picture = new Xdr.Picture();
            
            // Non-visual picture properties
            var nvPicPr = new Xdr.NonVisualPictureProperties();
            var cNvPr = new Xdr.NonVisualDrawingProperties() { Id = drawingId, Name = "Picture " + drawingId };
            var cNvPicPr = new Xdr.NonVisualPictureDrawingProperties();
            nvPicPr.Append(cNvPr);
            nvPicPr.Append(cNvPicPr);
            picture.Append(nvPicPr);
            
            // Blip fill (image reference)
            var blipFill = new Xdr.BlipFill();
            var blip = new A.Blip() { Embed = drawingsPart.GetIdOfPart(imagePart) };
            var stretch = new A.Stretch();
            stretch.Append(new A.FillRectangle());
            blipFill.Append(blip);
            blipFill.Append(stretch);
            picture.Append(blipFill);
            
            // Shape properties
            var spPr = new Xdr.ShapeProperties();
            var xfrm = new A.Transform2D();
            xfrm.Append(new A.Offset() { X = 0, Y = 0 });
            xfrm.Append(new A.Extents() { Cx = imageWidthEmu, Cy = imageHeightEmu });
            spPr.Append(xfrm);
            spPr.Append(new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle });
            picture.Append(spPr);
            
            // Add the picture to the anchor
            twoCellAnchor.Append(picture);
            
            // Add the client data
            twoCellAnchor.Append(new Xdr.ClientData());
            
            // Add the anchor to the drawing
            drawingsPart.WorksheetDrawing.Append(twoCellAnchor);
            
            // Save the drawing
            drawingsPart.WorksheetDrawing.Save();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error inserting image: {ex.Message}");
        }
    }
    
    private ImagePartType GetImagePartTypeFromFormat(string imagePath)
    {
        string extension = Path.GetExtension(imagePath).ToLower();
        
        switch (extension)
        {
            case ".jpg":
            case ".jpeg":
                return ImagePartType.Jpeg;
            case ".png":
                return ImagePartType.Png;
            case ".gif":
                return ImagePartType.Gif;
            case ".bmp":
                return ImagePartType.Bmp;
            case ".tiff":
            case ".tif":
                return ImagePartType.Tiff;
            default:
                return ImagePartType.Jpeg; // Default to JPEG
        }
    }

    private void ProcessForLoopPlaceholders(WorksheetPart worksheetPart)
    {
        if (worksheetPart?.Worksheet == null || _data.Collections == null || !_data.Collections.Any()) return;
        
        // Get all cells in the worksheet
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        
        // Get shared string table
        SharedStringTablePart sharedStringTablePart = worksheetPart.GetParentParts()
            .OfType<WorkbookPart>().FirstOrDefault()?.SharedStringTablePart;
        
        if (sharedStringTablePart == null) return;
        
        // Find for loop placeholders in the worksheet
        Dictionary<(int StartRow, int EndRow), (string CollectionName, List<string> TemplateRows)> forLoops = new Dictionary<(int, int), (string, List<string>)>();
        
        // First pass: Find for loop start and end markers
        Dictionary<int, (string CollectionName, int StartRow)> openForLoops = new Dictionary<int, (string, int)>();
        
        foreach (var row in sheetData.Elements<Row>())
        {
            // Check if this row contains a for loop start or end marker
            foreach (var cell in row.Elements<Cell>())
            {
                string cellText = GetCellText(cell, sharedStringTablePart);
                if (string.IsNullOrEmpty(cellText)) continue;
                
                // Check for for loop start marker: {% for item in items %}
                var startMatch = Regex.Match(cellText, @"\{%\s*for\s+(\w+)\s+in\s+(\w+)\s*%\}");
                if (startMatch.Success)
                {
                    string itemName = startMatch.Groups[1].Value;
                    string collectionName = startMatch.Groups[2].Value;
                    
                    // Check if this collection exists in our data
                    if (_data.HasCollection(collectionName))
                    {
                        // Store the start of the for loop
                        int rowIndex = (int)row.RowIndex.Value;
                        openForLoops[rowIndex] = (collectionName, rowIndex);
                    }
                    
                    break;
                }
                
                // Check for for loop end marker: {% endfor %}
                var endMatch = Regex.Match(cellText, @"\{%\s*endfor\s*%\}");
                if (endMatch.Success)
                {
                    // Find the matching start marker
                    if (openForLoops.Any())
                    {
                        var lastOpenLoop = openForLoops.OrderByDescending(kv => kv.Key).First();
                        int startRowIndex = lastOpenLoop.Value.StartRow;
                        string collectionName = lastOpenLoop.Value.CollectionName;
                        int endRowIndex = (int)row.RowIndex.Value;
                        
                        // Store the for loop range
                        var templateRows = ExtractTemplateRows(sheetData, startRowIndex + 1, endRowIndex - 1, sharedStringTablePart);
                        forLoops[(startRowIndex, endRowIndex)] = (collectionName, templateRows);
                        
                        // Remove the processed for loop from open loops
                        openForLoops.Remove(lastOpenLoop.Key);
                    }
                    
                    break;
                }
            }
        }
        
        // Process each for loop
        foreach (var forLoop in forLoops.OrderByDescending(kv => kv.Key.StartRow))
        {
            (int startRowIndex, int endRowIndex) = forLoop.Key;
            (string collectionName, List<string> templateRows) = forLoop.Value;
            
            // Get the collection data
            var collection = _data.GetCollection(collectionName);
            if (collection == null || !collection.Any())
            {
                // If collection is empty, remove the for loop rows
                RemoveRows(sheetData, startRowIndex, endRowIndex);
                continue;
            }
            
            // Remove the for loop marker rows
            RemoveRow(sheetData, endRowIndex);
            RemoveRow(sheetData, startRowIndex);
            
            // Insert the template rows for each item in the collection
            int currentRowIndex = startRowIndex;
            foreach (var item in collection)
            {
                foreach (var templateRow in templateRows)
                {
                    // Replace item properties in the template row
                    string processedRow = templateRow;
                    foreach (var property in item)
                    {
                        string placeholder = $"{{{{item.{property.Key}}}}}";
                        processedRow = processedRow.Replace(placeholder, property.Value ?? string.Empty);
                    }
                    
                    // Insert the processed row
                    InsertRowWithContent(sheetData, currentRowIndex, processedRow, sharedStringTablePart);
                    currentRowIndex++;
                }
            }
        }
        
        // Save the worksheet
        worksheetPart.Worksheet.Save();
    }
    
    private List<string> ExtractTemplateRows(SheetData sheetData, int startRowIndex, int endRowIndex, SharedStringTablePart sharedStringTablePart)
    {
        List<string> templateRows = new List<string>();
        
        for (int rowIndex = startRowIndex; rowIndex <= endRowIndex; rowIndex++)
        {
            var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowIndex);
            if (row == null) continue;
            
            // Extract the row content
            string rowContent = string.Empty;
            foreach (var cell in row.Elements<Cell>())
            {
                string cellText = GetCellText(cell, sharedStringTablePart);
                rowContent += $"{cell.CellReference.Value}:{cellText}|";
            }
            
            templateRows.Add(rowContent);
        }
        
        return templateRows;
    }
    
    private void InsertRowWithContent(SheetData sheetData, int rowIndex, string rowContent, SharedStringTablePart sharedStringTablePart)
    {
        // Create new row
        Row row = new Row() { RowIndex = (uint)rowIndex };
        
        // Parse the row content and create cells
        string[] cellContents = rowContent.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
        foreach (string cellContent in cellContents)
        {
            string[] parts = cellContent.Split(new char[] { ':' }, 2);
            if (parts.Length != 2) continue;
            
            string cellReference = parts[0];
            string cellText = parts[1];
            
            // Create cell
            Cell cell = new Cell() { CellReference = cellReference };
            
            // Set cell value
            if (!string.IsNullOrEmpty(cellText))
            {
                // Create a new shared string
                var sharedString = new SharedStringItem(new Text(cellText));
                sharedStringTablePart.SharedStringTable.AppendChild(sharedString);
                int index = sharedStringTablePart.SharedStringTable.Count() - 1;
                
                cell.DataType = CellValues.SharedString;
                cell.CellValue = new CellValue(index.ToString());
            }
            
            // Add cell to row
            row.AppendChild(cell);
        }
        
        // Insert row in the correct position
        Row nextRow = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value > rowIndex);
        if (nextRow != null)
        {
            sheetData.InsertBefore(row, nextRow);
        }
        else
        {
            sheetData.AppendChild(row);
        }
    }
    
    private void RemoveRows(SheetData sheetData, int startRowIndex, int endRowIndex)
    {
        for (int rowIndex = endRowIndex; rowIndex >= startRowIndex; rowIndex--)
        {
            RemoveRow(sheetData, rowIndex);
        }
    }
    
    private void RemoveRow(SheetData sheetData, int rowIndex)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.Value == rowIndex);
        if (row != null)
        {
            row.Remove();
        }
    }
}