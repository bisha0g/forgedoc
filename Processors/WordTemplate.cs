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
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Xml;
using HtmlAgilityPack;

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
                    
                    // Process image placeholders in main document
                    ProcessImagePlaceholders(doc.MainDocumentPart);

                    // Replace placeholders in headers
                    if (doc.MainDocumentPart.HeaderParts != null)
                    {
                        foreach (var headerPart in doc.MainDocumentPart.HeaderParts)
                        {
                            ReplacePlaceholdersInPart(headerPart);
                            ProcessTablePlaceholders(headerPart);
                            ProcessImagePlaceholders(headerPart);
                        }
                    }

                    // Replace placeholders in footers
                    if (doc.MainDocumentPart.FooterParts != null)
                    {
                        foreach (var footerPart in doc.MainDocumentPart.FooterParts)
                        {
                            ReplacePlaceholdersInPart(footerPart);
                            ProcessTablePlaceholders(footerPart);
                            ProcessImagePlaceholders(footerPart);
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
                    // Check if the placeholder value contains HTML
                    if (placeholder.Value != null && IsHtml(placeholder.Value))
                    {
                        // Mark that we found a placeholder, but don't replace it yet
                        // We'll handle HTML separately
                        containsPlaceholder = true;
                    }
                    else
                    {
                        // Regular text replacement
                        modifiedText = modifiedText.Replace(key, placeholder.Value ?? string.Empty);
                        containsPlaceholder = true;
                    }
                }
            }

            // If we found and replaced any placeholders, update the paragraph
            if (containsPlaceholder)
            {
                // Clear existing runs
                paragraph.RemoveAllChildren<Run>();

                // Process the text for each placeholder, handling HTML content
                foreach (var placeholder in _data.Placeholders)
                {
                    string key = $"{{{{{placeholder.Key}}}}}";
                    if (modifiedText.Contains(key))
                    {
                        if (placeholder.Value != null && IsHtml(placeholder.Value))
                        {
                            // Split the text at the placeholder
                            int placeholderIndex = modifiedText.IndexOf(key);
                            string beforePlaceholder = modifiedText.Substring(0, placeholderIndex);
                            string afterPlaceholder = modifiedText.Substring(placeholderIndex + key.Length);

                            // Add text before the placeholder
                            if (!string.IsNullOrEmpty(beforePlaceholder))
                            {
                                paragraph.AppendChild(new Run(new Text(beforePlaceholder)));
                            }

                            // Add the HTML content
                            AppendHtmlToRun(paragraph, placeholder.Value, null);

                            // Update the modified text to continue processing
                            modifiedText = afterPlaceholder;
                        }
                        else
                        {
                            // Regular replacement already handled above
                        }
                    }
                }

                // Add any remaining text
                if (!string.IsNullOrEmpty(modifiedText))
                {
                    paragraph.AppendChild(new Run(new Text(modifiedText)));
                }
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
                        bool hasHtmlPlaceholder = false;
                        
                        // First pass: check if there are any HTML placeholders
                        foreach (var placeholder in _data.Placeholders)
                        {
                            string key = $"{{{{{placeholder.Key}}}}}";
                            if (textModified.Contains(key) && placeholder.Value != null && IsHtml(placeholder.Value))
                            {
                                hasHtmlPlaceholder = true;
                                break;
                            }
                        }
                        
                        if (hasHtmlPlaceholder)
                        {
                            // If there's HTML content, we need to handle the entire run differently
                            var parentRun = text.Parent;
                            if (parentRun != null)
                            {
                                var runProperties = parentRun.Elements<RunProperties>().FirstOrDefault()?.CloneNode(true);
                                
                                // Get the text and process each placeholder
                                string runText = originalText;
                                foreach (var placeholder in _data.Placeholders)
                                {
                                    string key = $"{{{{{placeholder.Key}}}}}";
                                    if (runText.Contains(key))
                                    {
                                        // Split at the placeholder
                                        int placeholderIndex = runText.IndexOf(key);
                                        string beforePlaceholder = runText.Substring(0, placeholderIndex);
                                        string afterPlaceholder = runText.Substring(placeholderIndex + key.Length);
                                        
                                        // Add text before placeholder
                                        if (!string.IsNullOrEmpty(beforePlaceholder))
                                        {
                                            var newRun = new Run();
                                            if (runProperties != null)
                                                newRun.AppendChild(runProperties.CloneNode(true));
                                            newRun.AppendChild(new Text(beforePlaceholder));
                                            parentRun.InsertBeforeSelf(newRun);
                                        }
                                        
                                        // Add HTML content
                                        AppendHtmlToRun(parentRun.Parent, placeholder.Value, null);
                                        
                                        // Update text for next iteration
                                        runText = afterPlaceholder;
                                    }
                                }
                                
                                // Add any remaining text
                                if (!string.IsNullOrEmpty(runText))
                                {
                                    var newRun = new Run();
                                    if (runProperties != null)
                                        newRun.AppendChild(runProperties.CloneNode(true));
                                    newRun.AppendChild(new Text(runText));
                                    parentRun.InsertBeforeSelf(newRun);
                                }
                                
                                // Remove the original run
                                parentRun.Remove();
                            }
                        }
                        else
                        {
                            // Regular text replacement
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
    }
    
    private bool IsHtml(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
            return false;
            
        // Check for common HTML tags
        return text.Contains("<p") || text.Contains("<div") || text.Contains("<span") || 
               text.Contains("<br") || text.Contains("<b>") || text.Contains("<i>") || 
               text.Contains("<u>") || text.Contains("<strong") || text.Contains("<em") ||
               text.Contains("</p>") || text.Contains("</div>") || text.Contains("</span>") || 
               text.Contains("</b>") || text.Contains("</i>") || text.Contains("</u>") ||
               text.Contains("</strong>") || text.Contains("</em>");
    }
    
    private void AppendHtmlToParagraph(OpenXmlElement parentElement, string htmlContent)
    {
        try
        {
            // Load the HTML content
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);
            
            // Process the HTML nodes
            foreach (var node in htmlDoc.DocumentNode.ChildNodes)
            {
                ProcessHtmlNode(parentElement, node, null);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing HTML: {ex.Message}");
            // Fallback: add as plain text
            var run = new Run();
            run.AppendChild(new Text(htmlContent));
            parentElement.AppendChild(run);
        }
    }
    
    private void AppendHtmlToRun(OpenXmlElement parentElement, string htmlContent, RunProperties baseProperties = null)
    {
        try
        {
            // Load the HTML content
            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);
            
            // Process the HTML nodes
            foreach (var node in htmlDoc.DocumentNode.ChildNodes)
            {
                var run = new Run();
                if (baseProperties != null)
                    run.AppendChild(baseProperties.CloneNode(true));
                    
                ProcessHtmlNode(parentElement, node, baseProperties);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing HTML: {ex.Message}");
            // Fallback: add as plain text
            var run = new Run();
            if (baseProperties != null)
                run.AppendChild(baseProperties.CloneNode(true));
            run.AppendChild(new Text(htmlContent));
            parentElement.AppendChild(run);
        }
    }
    
    private void ProcessHtmlNode(OpenXmlElement parent, HtmlNode node, RunProperties baseProperties)
    {
        if (node.NodeType == HtmlNodeType.Text)
        {
            // Create a run with the text content
            var run = new Run();
            if (baseProperties != null)
                run.AppendChild(baseProperties.CloneNode(true));
                
            // Decode HTML entities
            string textContent = System.Net.WebUtility.HtmlDecode(node.InnerText);
            
            // Check if the text contains RTL characters
            bool isRtlText = ContainsRtlText(textContent);
            
            // If RTL text, add RTL properties to the run
            if (isRtlText)
            {
                var runProps = run.GetFirstChild<RunProperties>();
                if (runProps == null)
                {
                    runProps = new RunProperties();
                    run.PrependChild(runProps);
                }
                
                runProps.AppendChild(new RightToLeftText() { Val = OnOffValue.FromBoolean(true) });
            }
            
            run.AppendChild(new Text(textContent) { Space = SpaceProcessingModeValues.Preserve });
            
            // Add to parent
            parent.AppendChild(run);
            return;
        }
        
        // Process different HTML elements
        switch (node.Name.ToLower())
        {
            case "p":
                // For paragraphs in a table cell, we need to handle them differently
                if (parent is TableCell)
                {
                    var paragraph = new Paragraph();
                    
                    // Apply paragraph properties based on style attributes
                    var pPr = new ParagraphProperties();
                    
                    // Handle text alignment
                    if (node.Attributes["style"] != null)
                    {
                        string style = node.Attributes["style"].Value;
                        if (style.Contains("text-align: right") || style.Contains("text-align:right"))
                        {
                            pPr.AppendChild(new Justification() { Val = JustificationValues.Right });
                        }
                        else if (style.Contains("text-align: center") || style.Contains("text-align:center"))
                        {
                            pPr.AppendChild(new Justification() { Val = JustificationValues.Center });
                        }
                        else if (style.Contains("text-align: justify") || style.Contains("text-align:justify"))
                        {
                            pPr.AppendChild(new Justification() { Val = JustificationValues.Both });
                        }
                        
                        // Handle RTL text direction
                        if (style.Contains("direction: rtl") || style.Contains("direction:rtl"))
                        {
                            pPr.AppendChild(new BiDi() { Val = OnOffValue.FromBoolean(true) });
                        }
                    }
                    
                    // Check if the text contains RTL characters (Arabic, Hebrew, etc.)
                    bool containsRtlText = ContainsRtlText(node.InnerText);
                    if (containsRtlText)
                    {
                        pPr.AppendChild(new BiDi() { Val = OnOffValue.FromBoolean(true) });
                    }
                    
                    paragraph.AppendChild(pPr);
                    
                    // Process child nodes
                    foreach (var childNode in node.ChildNodes)
                    {
                        ProcessHtmlNode(paragraph, childNode, baseProperties);
                    }
                    
                    parent.AppendChild(paragraph);
                }
                else if (parent is Paragraph)
                {
                    // If we're already in a paragraph, just process the children
                    foreach (var childNode in node.ChildNodes)
                    {
                        ProcessHtmlNode(parent, childNode, baseProperties);
                    }
                }
                break;
                
            case "br":
                // Add a line break
                var breakRun = new Run();
                if (baseProperties != null)
                    breakRun.AppendChild(baseProperties.CloneNode(true));
                breakRun.AppendChild(new Break());
                parent.AppendChild(breakRun);
                break;
                
            case "b":
            case "strong":
                // Bold text
                foreach (var childNode in node.ChildNodes)
                {
                    var boldRun = new Run();
                    var boldProps = baseProperties != null ? baseProperties.CloneNode(true) as RunProperties : new RunProperties();
                    boldProps.AppendChild(new Bold());
                    boldRun.AppendChild(boldProps);
                    
                    if (childNode.NodeType == HtmlNodeType.Text)
                    {
                        boldRun.AppendChild(new Text(System.Net.WebUtility.HtmlDecode(childNode.InnerText)));
                        parent.AppendChild(boldRun);
                    }
                    else
                    {
                        ProcessHtmlNode(parent, childNode, boldProps);
                    }
                }
                break;
                
            case "i":
            case "em":
                // Italic text
                foreach (var childNode in node.ChildNodes)
                {
                    var italicRun = new Run();
                    var italicProps = baseProperties != null ? baseProperties.CloneNode(true) as RunProperties : new RunProperties();
                    italicProps.AppendChild(new Italic());
                    italicRun.AppendChild(italicProps);
                    
                    if (childNode.NodeType == HtmlNodeType.Text)
                    {
                        italicRun.AppendChild(new Text(System.Net.WebUtility.HtmlDecode(childNode.InnerText)));
                        parent.AppendChild(italicRun);
                    }
                    else
                    {
                        ProcessHtmlNode(parent, childNode, italicProps);
                    }
                }
                break;
                
            case "u":
                // Underlined text
                foreach (var childNode in node.ChildNodes)
                {
                    var underlineRun = new Run();
                    var underlineProps = baseProperties != null ? baseProperties.CloneNode(true) as RunProperties : new RunProperties();
                    underlineProps.AppendChild(new Underline() { Val = UnderlineValues.Single });
                    underlineRun.AppendChild(underlineProps);
                    
                    if (childNode.NodeType == HtmlNodeType.Text)
                    {
                        underlineRun.AppendChild(new Text(System.Net.WebUtility.HtmlDecode(childNode.InnerText)));
                        parent.AppendChild(underlineRun);
                    }
                    else
                    {
                        ProcessHtmlNode(parent, childNode, underlineProps);
                    }
                }
                break;
                
            case "span":
                // Handle span with style attributes
                var spanProps = baseProperties != null ? baseProperties.CloneNode(true) as RunProperties : new RunProperties();
                
                if (node.Attributes["style"] != null)
                {
                    string style = node.Attributes["style"].Value;
                    
                    // Handle text color
                    var colorMatch = Regex.Match(style, @"color:\s*#([0-9A-Fa-f]{6})");
                    if (colorMatch.Success)
                    {
                        string colorHex = colorMatch.Groups[1].Value;
                        spanProps.AppendChild(new DocumentFormat.OpenXml.Wordprocessing.Color() { Val = colorHex });
                    }
                    
                    // Handle font size
                    var fontSizeMatch = Regex.Match(style, @"font-size:\s*(\d+)pt");
                    if (fontSizeMatch.Success)
                    {
                        int fontSize = int.Parse(fontSizeMatch.Groups[1].Value);
                        spanProps.AppendChild(new FontSize() { Val = (fontSize * 2).ToString() }); // Convert pt to half-points
                    }
                    
                    // Handle font family
                    var fontFamilyMatch = Regex.Match(style, @"font-family:\s*([^;]+)");
                    if (fontFamilyMatch.Success)
                    {
                        string fontFamily = fontFamilyMatch.Groups[1].Value.Trim().Trim('\'', '"');
                        spanProps.AppendChild(new RunFonts() { Ascii = fontFamily, HighAnsi = fontFamily });
                    }
                }
                
                foreach (var childNode in node.ChildNodes)
                {
                    if (childNode.NodeType == HtmlNodeType.Text)
                    {
                        var spanRun = new Run();
                        spanRun.AppendChild(spanProps.CloneNode(true));
                        spanRun.AppendChild(new Text(System.Net.WebUtility.HtmlDecode(childNode.InnerText)));
                        parent.AppendChild(spanRun);
                    }
                    else
                    {
                        ProcessHtmlNode(parent, childNode, spanProps);
                    }
                }
                break;
                
            case "div":
                // For divs, process children
                foreach (var childNode in node.ChildNodes)
                {
                    ProcessHtmlNode(parent, childNode, baseProperties);
                }
                break;
                
            default:
                // For other elements, just process the inner text
                if (!string.IsNullOrWhiteSpace(node.InnerText))
                {
                    var defaultRun = new Run();
                    if (baseProperties != null)
                        defaultRun.AppendChild(baseProperties.CloneNode(true));
                    defaultRun.AppendChild(new Text(System.Net.WebUtility.HtmlDecode(node.InnerText)));
                    parent.AppendChild(defaultRun);
                }
                break;
        }
    }
    
    private bool ContainsRtlText(string text)
    {
        // Check for RTL characters (Arabic, Hebrew, etc.)
        return text.Any(c => c >= 0x0590 && c <= 0x05FF || c >= 0x0600 && c <= 0x06FF || c >= 0xFB50 && c <= 0xFDFF || c >= 0xFE70 && c <= 0xFEFF);
    }
    
    private void ProcessImagePlaceholders(OpenXmlPart part)
    {
        if (part?.RootElement == null || _data.Images == null || !_data.Images.Any()) return;
        
        Console.WriteLine($"Processing image placeholders in {part.GetType().Name}");
        
        // Get all paragraphs in the document
        var paragraphs = part.RootElement.Descendants<Paragraph>().ToList();
        
        foreach (var paragraph in paragraphs)
        {
            // Get the text of the paragraph
            string paragraphText = GetParagraphText(paragraph);
            
            // Skip if no text
            if (string.IsNullOrWhiteSpace(paragraphText)) continue;
            
            Console.WriteLine($"Checking paragraph: '{paragraphText}'");
            
            // List to store image placeholders found in this paragraph
            List<string> imagePlaceholders = new List<string>();
            
            // Find {% key %} or {% key:widthxheight %} style placeholders
            // Updated regex to match the format {% SupervisorSignature:200x100 %}
            var matches = Regex.Matches(paragraphText, @"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
            foreach (Match match in matches)
            {
                imagePlaceholders.Add(match.Value);
                Console.WriteLine($"Found image placeholder: {match.Value}");
            }
            
            // Also check for any text that might include dimensions but not in the expected format
            var dimensionMatches = Regex.Matches(paragraphText, @":\d+x\d+");
            foreach (Match match in dimensionMatches)
            {
                // Find the surrounding placeholder-like text
                int startIndex = paragraphText.LastIndexOf("{%", match.Index);
                int endIndex = paragraphText.IndexOf("%}", match.Index);
                
                if (startIndex >= 0 && endIndex > startIndex)
                {
                    string fullPlaceholder = paragraphText.Substring(startIndex, endIndex - startIndex + 2);
                    if (!imagePlaceholders.Contains(fullPlaceholder))
                    {
                        imagePlaceholders.Add(fullPlaceholder);
                        Console.WriteLine($"Found dimension-containing text: {fullPlaceholder}");
                    }
                }
            }
            
            // Also check for standalone dimension patterns like "100x100"
            var standaloneDimensions = Regex.Matches(paragraphText, @"\b\d+x\d+\b");
            foreach (Match match in standaloneDimensions)
            {
                // Look for nearby placeholder markers
                int beforeIndex = Math.Max(0, match.Index - 20);
                int afterIndex = Math.Min(paragraphText.Length - 1, match.Index + match.Length + 20);
                string surrounding = paragraphText.Substring(beforeIndex, afterIndex - beforeIndex);
                
                if (surrounding.Contains("{%") && surrounding.Contains("%}"))
                {
                    // Try to extract the full placeholder
                    int startIndex = surrounding.IndexOf("{%");
                    if (startIndex >= 0)
                    {
                        startIndex += beforeIndex;
                        int endIndex = paragraphText.IndexOf("%}", startIndex);
                        if (endIndex > startIndex)
                        {
                            string fullPlaceholder = paragraphText.Substring(startIndex, endIndex - startIndex + 2);
                            if (!imagePlaceholders.Contains(fullPlaceholder))
                            {
                                imagePlaceholders.Add(fullPlaceholder);
                                Console.WriteLine($"Found placeholder with nearby dimensions: {fullPlaceholder}");
                            }
                        }
                    }
                }
            }
            
            // Process each placeholder
            foreach (string placeholder in imagePlaceholders)
            {
                string key = null;
                
                // Extract the key from the placeholder
                if (placeholder.StartsWith("{%"))
                {
                    // Extract key from {% key %} or {% key:widthxheight %}
                    var match = Regex.Match(placeholder, @"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
                    if (!match.Success)
                    {
                        // Try alternate format with spaces
                        match = Regex.Match(placeholder, @"\{%\s*([^:}]+)\s*:\s*(\d+)\s*x\s*(\d+)\s*%\}");
                    }
                    
                    if (match.Success)
                    {
                        key = match.Groups[1].Value.Trim();
                        Console.WriteLine($"Extracted key from placeholder: '{key}'");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to extract key from placeholder: {placeholder}");
                        continue;
                    }
                }
                
                // Skip if no key found
                if (string.IsNullOrEmpty(key))
                {
                    Console.WriteLine("No key found in placeholder");
                    continue;
                }
                
                // Get the image path from the data
                string imagePath = _data.GetImage(key);
                if (string.IsNullOrEmpty(imagePath))
                {
                    Console.WriteLine($"No image found for key: {key}");
                    continue;
                }
                
                // Check if the image file exists
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    continue;
                }
                
                // Insert the image
                InsertImageInParagraph(part, paragraph, imagePath, placeholder);
            }
        }
    }
    
    private void InsertImageInParagraph(OpenXmlPart part, Paragraph paragraph, string imagePath, string placeholderText)
    {
        try
        {
            Console.WriteLine($"Starting image insertion for placeholder: {placeholderText}, image path: {imagePath}");
            
            // Get the MainDocumentPart
            MainDocumentPart mainPart = part as MainDocumentPart;
            if (mainPart == null && part is HeaderPart headerPart)
            {
                mainPart = headerPart.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
                Console.WriteLine("Getting MainDocumentPart from HeaderPart");
            }
            else if (mainPart == null && part is FooterPart footerPart)
            {
                mainPart = footerPart.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
                Console.WriteLine("Getting MainDocumentPart from FooterPart");
            }
            
            if (mainPart == null)
            {
                mainPart = part.GetParentParts().OfType<MainDocumentPart>().FirstOrDefault();
                Console.WriteLine("Getting MainDocumentPart from parent parts");
            }
            
            if (mainPart == null)
            {
                Console.WriteLine("ERROR: Could not find MainDocumentPart");
                return;
            }
            
            // Get image dimensions
            int imageWidthEmu;
            int imageHeightEmu;
            ImagePartType imageType;
            
            using (var img = System.Drawing.Image.FromFile(imagePath))
            {
                // Check if we need to resize the image
                double maxWidthInPixels = 400; // Default max width in pixels
                double maxHeightInPixels = 300; // Default max height in pixels
                
                Console.WriteLine($"Checking for size in placeholder: {placeholderText}");
                
                // Try different regex patterns to match the size
                var sizeMatch = Regex.Match(placeholderText, @"\{%\s*([^:}]+):(\d+)x(\d+)\s*%\}");
                if (!sizeMatch.Success)
                {
                    // Try alternate format with spaces
                    sizeMatch = Regex.Match(placeholderText, @"\{%\s*([^:}]+)\s*:\s*(\d+)\s*x\s*(\d+)\s*%\}");
                }
                
                if (sizeMatch.Success && sizeMatch.Groups.Count > 2)
                {
                    Console.WriteLine($"Size match groups: {sizeMatch.Groups.Count}, Group 1: '{sizeMatch.Groups[1].Value}', Group 2: '{sizeMatch.Groups[2].Value}', Group 3: '{sizeMatch.Groups[3].Value}'");
                    
                    // Extract width and height from the placeholder
                    if (int.TryParse(sizeMatch.Groups[2].Value, out int width))
                    {
                        maxWidthInPixels = width;
                        Console.WriteLine($"Parsed width: {maxWidthInPixels}");
                    }
                    
                    if (int.TryParse(sizeMatch.Groups[3].Value, out int height))
                    {
                        maxHeightInPixels = height;
                        Console.WriteLine($"Parsed height: {maxHeightInPixels}");
                    }
                    
                    Console.WriteLine($"Found size in placeholder: {maxWidthInPixels}x{maxHeightInPixels}");
                }
                else
                {
                    Console.WriteLine($"No size information found in placeholder or could not parse: {placeholderText}");
                }
                
                // Calculate new dimensions while maintaining aspect ratio
                double scale = 1.0;
                if (img.Width > maxWidthInPixels || img.Height > maxHeightInPixels)
                {
                    double widthScale = maxWidthInPixels / img.Width;
                    double heightScale = maxHeightInPixels / img.Height;
                    scale = Math.Min(widthScale, heightScale);
                    
                    Console.WriteLine($"Resizing image with scale factor: {scale}");
                }
                
                int newWidth = (int)(img.Width * scale);
                int newHeight = (int)(img.Height * scale);
                
                // Convert pixels to EMUs (English Metric Units)
                // 1 inch = 914400 EMUs, 1 inch = 96 pixels (default)
                double emuPerPixel = 9525;
                imageWidthEmu = (int)(newWidth * emuPerPixel);
                imageHeightEmu = (int)(newHeight * emuPerPixel);
                Console.WriteLine($"Original dimensions: {img.Width}x{img.Height} pixels");
                Console.WriteLine($"New dimensions: {newWidth}x{newHeight} pixels, {imageWidthEmu}x{imageHeightEmu} EMUs");
                
                // Determine image type based on format
                imageType = GetImagePartTypeFromFormat(img.RawFormat);
                Console.WriteLine($"Detected image format: {imageType}");
            }
            
            // Add the image to the document
            ImagePart imagePart = mainPart.AddImagePart(imageType);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
                Console.WriteLine("Image data fed to ImagePart");
            }
            
            // Create the drawing element
            Drawing drawing = AddImageToAnyWhere(mainPart.Document, mainPart.GetIdOfPart(imagePart), imageWidthEmu, imageHeightEmu, imageType);
            Console.WriteLine("Drawing element created");
            
            // Replace the placeholder text with the image
            ReplaceTextWithImage(paragraph, placeholderText, drawing);
            Console.WriteLine("Placeholder text replaced with image");
        }
        catch (Exception ex)
        {
            // Log the error or handle it as appropriate
            Console.WriteLine($"ERROR inserting image: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
    
    private void ReplaceTextWithImage(Paragraph paragraph, string placeholderText, Drawing drawing)
    {
        Console.WriteLine($"Starting ReplaceTextWithImage for placeholder: {placeholderText}");
        
        try
        {
            // Create a completely new paragraph to replace the original one
            Paragraph newParagraph = new Paragraph();
            
            // Copy paragraph properties if they exist
            if (paragraph.ParagraphProperties != null)
            {
                newParagraph.ParagraphProperties = (ParagraphProperties)paragraph.ParagraphProperties.CloneNode(true);
            }
            
            // Get the text content of the paragraph
            string paragraphText = GetParagraphText(paragraph);
            Console.WriteLine($"Full paragraph text: '{paragraphText}'");
            
            // Find the placeholder in the paragraph text
            int placeholderIndex = paragraphText.IndexOf(placeholderText);
            
            // If the exact placeholder isn't found, try to find a similar pattern
            if (placeholderIndex < 0)
            {
                // Look for patterns like "{% key:100x100 %}" or ":100x100 %}" or just "100x100 %}"
                var matches = Regex.Matches(paragraphText, @"\{%[^}]+%\}|\:\d+x\d+\s*%\}|\d+x\d+\s*%\}");
                foreach (Match match in matches)
                {
                    Console.WriteLine($"Found potential placeholder fragment: '{match.Value}'");
                    placeholderIndex = match.Index;
                    placeholderText = match.Value;
                    
                    // If we found a fragment like ":100x100 %}", try to find the start of the placeholder
                    if (placeholderText.StartsWith(":"))
                    {
                        int startIndex = paragraphText.LastIndexOf("{%", placeholderIndex);
                        if (startIndex >= 0)
                        {
                            placeholderIndex = startIndex;
                            placeholderText = paragraphText.Substring(startIndex, match.Index + match.Length - startIndex);
                            Console.WriteLine($"Expanded placeholder to: '{placeholderText}'");
                        }
                    }
                    break;
                }
            }
            
            if (placeholderIndex < 0)
            {
                Console.WriteLine("Could not find placeholder in paragraph text");
                return;
            }
            
            // Create text runs for content before and after the placeholder
            if (placeholderIndex > 0)
            {
                string beforeText = paragraphText.Substring(0, placeholderIndex);
                newParagraph.AppendChild(new Run(new Text(beforeText)));
                Console.WriteLine($"Added text before placeholder: '{beforeText}'");
            }
            
            // Add the image
            newParagraph.AppendChild(new Run(drawing));
            Console.WriteLine("Added image to paragraph");
            
            // Add text after the placeholder
            int afterIndex = placeholderIndex + placeholderText.Length;
            if (afterIndex < paragraphText.Length)
            {
                string afterText = paragraphText.Substring(afterIndex);
                newParagraph.AppendChild(new Run(new Text(afterText)));
                Console.WriteLine($"Added text after placeholder: '{afterText}'");
            }
            
            // Replace the original paragraph with our new one
            paragraph.Parent.ReplaceChild(newParagraph, paragraph);
            Console.WriteLine("Replaced original paragraph with new paragraph containing the image");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in ReplaceTextWithImage: {ex.Message}");
            Console.WriteLine($"Stack trace: {ex.StackTrace}");
        }
    }
    
    private void ProcessCellImagePlaceholders(TableCell cell, Paragraph paragraph, Dictionary<string, string> dataItem = null)
    {
        if (_data.Images == null || !_data.Images.Any()) return;
        
        // Get the text of the paragraph
        string paragraphText = GetParagraphText(paragraph);
        
        // Skip if no text
        if (string.IsNullOrWhiteSpace(paragraphText)) return;
        
        // Find {% key %} or {% key:widthxheight %} style placeholders
        var matches = Regex.Matches(paragraphText, @"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
        
        foreach (Match match in matches)
        {
            string placeholder = match.Value;
            string key = match.Groups[1].Value.Trim();
            
            // If we have a dataItem and this is a Signature placeholder, use the SignatureKey
            if (dataItem != null && key == "Signature" && dataItem.ContainsKey("SignatureKey"))
            {
                key = dataItem["SignatureKey"];
            }
            
            // Get the image path from the data
            string imagePath = _data.GetImage(key);
            if (string.IsNullOrEmpty(imagePath))
            {
                Console.WriteLine($"No image found for key: {key}");
                continue;
            }
            
            // Check if the image file exists
            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Image file not found: {imagePath}");
                continue;
            }
            
            // Get the part containing the cell
            OpenXmlPart part = cell.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart;
            if (part == null)
            {
                // Try to get the part from the header or footer
                var header = cell.Ancestors<Header>().FirstOrDefault();
                if (header != null)
                {
                    part = header.HeaderPart;
                }
                else
                {
                    var footer = cell.Ancestors<Footer>().FirstOrDefault();
                    if (footer != null)
                    {
                        part = footer.FooterPart;
                    }
                }
            }
            
            if (part != null)
            {
                // Insert the image in the paragraph
                InsertImageInParagraph(part, paragraph, imagePath, placeholder);
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
                    
                    // Check for image placeholders {% ImageKey %} in the cell
                    var imagePlaceholderPattern = new Regex(@"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
                    var imageMatches = imagePlaceholderPattern.Matches(processedText);
                    
                    if (imageMatches.Count > 0)
                    {
                        // If we have image placeholders, we need to handle them specially
                        foreach (Match match in imageMatches)
                        {
                            string fullPlaceholder = match.Value;
                            string imageKey = match.Groups[1].Value.Trim();
                            
                            // Check if this is a SignatureKey reference from the data item
                            if (dataItem.ContainsKey("SignatureKey") && imageKey == "Signature")
                            {
                                // Replace the placeholder with the actual image key
                                string actualImageKey = dataItem["SignatureKey"];
                                
                                // Create a new placeholder with the actual key
                                string newPlaceholder = fullPlaceholder.Replace(imageKey, actualImageKey);
                                
                                // Replace in the processed text
                                processedText = processedText.Replace(fullPlaceholder, newPlaceholder);
                                
                                // Set the flag to indicate we made a replacement
                                replacementMade = true;
                            }
                        }
                    }
                    
                    // If we made any replacements, update the paragraph text
                    if (replacementMade)
                    {
                        // Final cleanup of any remaining table tags
                        processedText = Regex.Replace(processedText, @"\{\{#docTable\s+[^}]+\}\}", "");
                        processedText = processedText.Replace("{{/docTable}}", "");
                        processedText = processedText.Trim();
                        
                        // Update the paragraph with the new text
                        ReplaceParagraphText(paragraph, processedText);
                    }
                }
            }
            
            // Process image placeholders in the row after all text replacements are done
            foreach (var cell in newRow.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    ProcessCellImagePlaceholders(cell, paragraph, dataItem);
                }
            }
        }
    }
    
    private void CreateTableRows(Table parentTable, TableRow originalRow, List<Dictionary<string, string>> tableData)
    {
        // For each row in the table data
        for (int i = 0; i < tableData.Count; i++)
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
                        if (placeholder.StartsWith("item."))
                        {
                            string itemProperty = placeholder.Substring(5); // Remove "item." prefix
                            if (tableData[i].ContainsKey(itemProperty))
                            {
                                processedText = processedText.Replace($"{{{{{placeholder}}}}}", tableData[i][itemProperty]);
                                replacementMade = true;
                            }
                        }
                        // Check if this is a direct property name
                        else if (tableData[i].ContainsKey(placeholder))
                        {
                            processedText = processedText.Replace($"{{{{{placeholder}}}}}", tableData[i][placeholder]);
                            replacementMade = true;
                        }
                    }
                    
                    // Check for image placeholders {% ImageKey %} in the cell
                    var imagePlaceholderPattern = new Regex(@"\{%\s*([^:}]+)(?::(\d+)x(\d+))?\s*%\}");
                    var imageMatches = imagePlaceholderPattern.Matches(processedText);
                    
                    if (imageMatches.Count > 0)
                    {
                        // If we have image placeholders, we need to handle them specially
                        foreach (Match match in imageMatches)
                        {
                            string fullPlaceholder = match.Value;
                            string imageKey = match.Groups[1].Value.Trim();
                            
                            // Check if this is a SignatureKey reference from the data item
                            if (tableData[i].ContainsKey("SignatureKey") && imageKey == "Signature")
                            {
                                // Replace the placeholder with the actual image key
                                string actualImageKey = tableData[i]["SignatureKey"];
                                
                                // Create a new placeholder with the actual key
                                string newPlaceholder = fullPlaceholder.Replace(imageKey, actualImageKey);
                                
                                // Replace in the processed text
                                processedText = processedText.Replace(fullPlaceholder, newPlaceholder);
                                
                                // Set the flag to indicate we made a replacement
                                replacementMade = true;
                            }
                        }
                    }
                    
                    // If we made any replacements, update the paragraph text
                    if (replacementMade)
                    {
                        // Final cleanup of any remaining table tags
                        processedText = Regex.Replace(processedText, @"\{\{#docTable\s+[^}]+\}\}", "");
                        processedText = processedText.Replace("{{/docTable}}", "");
                        processedText = processedText.Trim();
                        
                        // Update the paragraph with the new text
                        ReplaceParagraphText(paragraph, processedText);
                    }
                }
            }
            
            // Process image placeholders in the row after all text replacements are done
            foreach (var cell in newRow.Elements<TableCell>())
            {
                foreach (var paragraph in cell.Elements<Paragraph>())
                {
                    ProcessCellImagePlaceholders(cell, paragraph, tableData[i]);
                }
            }
            
            // Add the new row to the table after the last row we added
            parentTable.InsertAfter(newRow, originalRow);
            originalRow = newRow;
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
            
            // Add the new row to the table after the last row we added
            parentTable.InsertAfter(newRow, originalRow);
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
            
            // Check if any values contain HTML
            bool hasHtmlContent = rowData.Values.Any(v => v != null && IsHtml(v));
            
            if (hasHtmlContent)
            {
                // Create text with all the data, handling HTML content
                foreach (var kv in rowData)
                {
                    // Add the key
                    Run keyRun = new Run();
                    keyRun.AppendChild(new Text($"{kv.Key}: "));
                    paragraph.AppendChild(keyRun);
                    
                    // Add the value, handling HTML if needed
                    if (kv.Value != null && IsHtml(kv.Value))
                    {
                        AppendHtmlToParagraph(paragraph, kv.Value);
                    }
                    else
                    {
                        Run valueRun = new Run();
                        valueRun.AppendChild(new Text(kv.Value ?? string.Empty));
                        paragraph.AppendChild(valueRun);
                    }
                    
                    // Add separator between key-value pairs
                    if (kv.Key != rowData.Keys.Last())
                    {
                        Run separatorRun = new Run();
                        separatorRun.AppendChild(new Text(", "));
                        paragraph.AppendChild(separatorRun);
                    }
                }
            }
            else
            {
                // Create text with all the data (no HTML)
                Run run = new Run();
                string text = string.Join(", ", rowData.Select(kv => $"{kv.Key}: {kv.Value}"));
                run.AppendChild(new Text(text));
                paragraph.AppendChild(run);
            }
            
            cell.AppendChild(paragraph);
        }
    }
    
    private void ReplaceParagraphText(Paragraph paragraph, string newText)
    {
        // Clear existing runs
        paragraph.RemoveAllChildren();
        
        // Check if the new text contains HTML
        if (IsHtml(newText))
        {
            // Process HTML content
            AppendHtmlToParagraph(paragraph, newText);
        }
        else
        {
            // Add a new run with the new text
            Run run = new Run();
            run.AppendChild(new Text(newText));
            paragraph.AppendChild(run);
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
    
    private Drawing AddImageToAnyWhere(Document mainPartDocument, string getIdOfPart, int imageSizeWidth, int imageSizeHeight, ImagePartType imageType = ImagePartType.Jpeg)
    {
        // Create a unique ID for the image
        string imageId = $"image{Guid.NewGuid()}";
        
        // Determine file extension based on image type
        string fileExtension = GetFileExtensionFromImageType(imageType);
        
        // Create a new Drawing object
        Drawing drawing = new Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = imageSizeWidth, Cy = imageSizeHeight },
                new DW.EffectExtent()
                {
                    LeftEdge = 0L,
                    TopEdge = 0L,
                    RightEdge = 0L,
                    BottomEdge = 0L
                },
                new DW.DocProperties()
                {
                    Id = (UInt32Value)1U,
                    Name = imageId
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks() { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties()
                                {
                                    Id = (UInt32Value)0U,
                                    Name = $"{imageId}{fileExtension}"
                                },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip(
                                    new A.BlipExtensionList(
                                        new A.BlipExtension()
                                        {
                                            Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                        })
                                )
                                {
                                    Embed = getIdOfPart,
                                    CompressionState = A.BlipCompressionValues.Print
                                },
                                new A.Stretch(
                                    new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = imageSizeWidth, Cy = imageSizeHeight }),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                )
                                { Preset = A.ShapeTypeValues.Rectangle }))
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
            
        return drawing;
    }
    
    private string GetFileExtensionFromImageType(ImagePartType imageType)
    {
        switch (imageType)
        {
            case ImagePartType.Jpeg:
                return ".jpg";
            case ImagePartType.Png:
                return ".png";
            case ImagePartType.Gif:
                return ".gif";
            case ImagePartType.Bmp:
                return ".bmp";
            case ImagePartType.Tiff:
                return ".tiff";
            default:
                return ".jpg";
        }
    }
    
    private ImagePartType GetImagePartTypeFromFormat(System.Drawing.Imaging.ImageFormat format)
    {
        if (format.Equals(System.Drawing.Imaging.ImageFormat.Jpeg))
            return ImagePartType.Jpeg;
        else if (format.Equals(System.Drawing.Imaging.ImageFormat.Png))
            return ImagePartType.Png;
        else if (format.Equals(System.Drawing.Imaging.ImageFormat.Gif))
            return ImagePartType.Gif;
        else if (format.Equals(System.Drawing.Imaging.ImageFormat.Bmp))
            return ImagePartType.Bmp;
        else if (format.Equals(System.Drawing.Imaging.ImageFormat.Tiff))
            return ImagePartType.Tiff;
        else
            return ImagePartType.Jpeg; // Default to JPEG
    }
}