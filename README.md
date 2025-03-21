# ForgeDoc - Word Template Processor

ForgeDoc is a library for processing Word templates with placeholders and generating dynamic Word documents.

## Features

- Replace text placeholders in Word documents
- Add images to Word documents using `{% imageKey %}` syntax
- Render tables in Word documents
- Support for rich text formatting
- Support for special characters with specific fonts (e.g., Wingdings)
- Works with headers, footers, and tables

## Image Placeholders

ForgeDoc supports inserting images into Word documents using a special placeholder syntax:

```text
{% imageKey %}
```

Where `imageKey` is a key that references an image path you've added to the template data.

### Image Resizing

You can specify custom dimensions for your images using the following syntax:

```text
{% imageKey:widthxheight %}
```

For example:

```text
{% Logo:200x100 %}
```

This will resize the image to a maximum width of 200 pixels and a maximum height of 100 pixels, while maintaining the aspect ratio. If you don't specify dimensions, the default maximum size is 400x300 pixels.

#### Important Notes on Image Resizing Syntax

- Make sure there are no spaces between the key and the colon: `{% Logo:200x100 %}` (correct) vs `{% Logo: 200x100 %}` (incorrect)
- Make sure there are no spaces in the dimensions: `200x100` (correct) vs `200 x 100` (incorrect)
- The full syntax should look exactly like: `{% SupervisorSignature:200x100 %}`


### Adding Images to Template Data

```csharp
// Create template data
var data = new WordTemplateData();

// Add image (path must be accessible at runtime)
data.AddImage("Logo", @"C:\path\to\logo.png");
data.AddImage("Signature", @"C:\path\to\signature.jpg");

// Process the template
var template = new WordTemplate("template.docx", data);
template.Process("output.docx");
```

### Images in Tables

You can also include images in tables by using the same placeholder syntax:

```csharp
var tableData = new List<Dictionary<string, string>>
{
    new Dictionary<string, string> { { "Name", "John Doe" }, { "SignatureKey", "Signature1" } },
    new Dictionary<string, string> { { "Name", "Jane Smith" }, { "SignatureKey", "Signature2" } }
};

// Add the table data
data.AddTable("Employees", tableData);

// Add the images
data.AddImage("Signature1", @"C:\path\to\john_signature.png");
data.AddImage("Signature2", @"C:\path\to\jane_signature.png");
```

In your Word template, use:
```text
{{#docTable Employees}}
Name: {{ item.Name }}
Signature: {% {{ item.SignatureKey }} %}
{{/docTable}}
```

### Working with Database Images

When working with images from a database, save them to temporary files first:

```csharp
// Save database image to a temporary file
string tempPath = Path.Combine(Path.GetTempPath(), $"signature_{Guid.NewGuid()}.png");
File.WriteAllBytes(tempPath, databaseImageBytes);

// Add the image to the template data
data.AddImage("Signature", tempPath);

// Remember to clean up temporary files after processing
try {
    if (File.Exists(tempPath)) {
        File.Delete(tempPath);
    }
} catch {
    // Handle cleanup errors
}
```

### Supported Image Formats

The processor supports common image formats:

  • PNG
  • JPEG/JPG
  • GIF
  • BMP
  • TIFF

Images are inserted at their original size.

## Table Rendering

ForgeDoc supports three ways to render tables in Word documents:

### 1. Standard Table Syntax

Use the following syntax in your Word template to define where a table should be rendered:

```text
{{#docTable tableName}}
{{item.column1}} {{item.column2}}
{{/docTable}}
```

This will create a new table at the location of the placeholder, with one row for each item in the table data.

### 2. Existing Table with Placeholders

You can also use placeholders directly in an existing table in your Word document:

```markdown
| Header 1   | Header 2   | Header 3   |
|------------|------------|------------|
| {{column1}} | {{column2}} | {{column3}} |
```

The processor will replace the placeholders with values from the first row of data and duplicate the row for each additional data item.

### 3. Mixed Syntax in Existing Tables

ForgeDoc also supports a mixed syntax approach where you can have both table placeholders and regular placeholders in the same table:

```markdown
| Header 1   | Header 2   | Header 3   |
|------------|------------|------------|

| {{#docTable tableName}}{{column1}} | {{column2}} | {{column3}}{{/docTable}} |

```

In this case, the processor will:

1. Identify the table that contains the `{{#docTable}}` tag
2. Replace all placeholders in the row with values from the data
3. Duplicate the row for each data item in the table data
4. Remove any remaining table tags

This is particularly useful for complex tables with headers in different languages or when working with existing template documents.

## Special Characters with Specific Fonts

ForgeDoc supports special characters with specific fonts in Word templates. This feature is particularly useful for inserting symbols like checkmarks, bullets, and other special characters that require specific fonts like Wingdings, Wingdings 2, Symbol, etc.

### Usage

#### In the Template

In your Word template, use the standard placeholder syntax with double curly braces:

```
{{CheckMark}}
```

#### In Your Code

When setting up your template data, use the `AddSpecialCharacter` method to specify the character and its font:

```csharp
var data = new WordTemplateData();

// Add a special character with a specific font
// \uf052 is a checkmark in Wingdings 2 font
data.AddSpecialCharacter("CheckMark", "\uf052", "Wingdings 2");
```

### Common Special Characters in Wingdings 2

Here are some common special characters in the Wingdings 2 font:

| Character | Unicode | Description |
|-----------|---------|-------------|
| \uf052    | U+F052  | Checkmark   |
| \uf06E    | U+F06E  | Circle      |
| \uf06F    | U+F06F  | Square      |
| \uf070    | U+F070  | Diamond     |
| \uf071    | U+F071  | Triangle    |
| \uf0FC    | U+F0FC  | Arrow Right |
| \uf0FB    | U+F0FB  | Arrow Left  |
| \uf0FC    | U+F0FC  | Arrow Up    |
| \uf0FD    | U+F0FD  | Arrow Down  |

### Example

```csharp
// Create template data
var data = new WordTemplateData();

// Add regular placeholders
data.AddPlaceholder("Title", "Special Character Example");

// Add special characters with specific fonts
data.AddSpecialCharacter("CheckMark", "\uf052", "Wingdings 2");
data.AddSpecialCharacter("Square", "\uf06F", "Wingdings 2");
data.AddSpecialCharacter("Circle", "\uf06E", "Wingdings 2");

// Create and process the template
var template = new WordTemplate("template.docx", data);
byte[] result = template.GetFile();
```

### Notes

- The character must be provided as a Unicode escape sequence (e.g., `\uf052`) or as the actual character.
- The font name must exactly match the font name in Word (e.g., "Wingdings 2", "Symbol", etc.).
- This feature works in all parts of the document, including headers, footers, and tables.

## Usage

### Basic Usage

```csharp
// Create a new WordTemplateData object
var data = new WordTemplateData
{
    Placeholders = new Dictionary<string, string>
    {
        { "Title", "My Document" },
        { "Author", "John Doe" },
        { "Date", DateTime.Now.ToString("yyyy-MM-dd") }
    }
};

// Create a new WordTemplate object
var template = new WordTemplate("path/to/template.docx", data);

// Get the generated document as a byte array
byte[] document = template.GetFile();
```

### Adding Images

```csharp
var data = new WordTemplateData
{
    Placeholders = new Dictionary<string, string>
    {
        { "Title", "My Document with Images" }
    },
    Images = new Dictionary<string, string>
    {
        { "Logo", "path/to/logo.png" },
        { "Signature", "path/to/signature.jpg" }
    }
};
```

In your Word template, use `{% Logo %}` and `{% Signature %}` to place the images.

### Adding Tables

Tables can be added to your Word template using the following syntax:

```text
{{#docTable tableName}}
{{item.column1}} {{item.column2}} {{item.column3}}
{{/docTable}}
```

In your code, add the table data like this:

```csharp
var tableData = new List<Dictionary<string, string>>
{
    new Dictionary<string, string>
    {
        { "column1", "Row 1, Column 1" },
        { "column2", "Row 1, Column 2" },
        { "column3", "Row 1, Column 3" }
    },
    new Dictionary<string, string>
    {
        { "column1", "Row 2, Column 1" },
        { "column2", "Row 2, Column 2" },
        { "column3", "Row 2, Column 3" }
    }
};

data.AddTable("tableName", tableData);
```

The table will be rendered with one row for each item in the list, and one column for each placeholder in the template.

## Example

Here's a complete example of using all features:

```csharp
var data = new WordTemplateData
{
    Placeholders = new Dictionary<string, string>
    {
        { "Title", "Inventory Report" },
        { "Date", DateTime.Now.ToString("yyyy-MM-dd") },
        { "PreparedBy", "John Doe" }
    },
    Images = new Dictionary<string, string>
    {
        { "CompanyLogo", "path/to/logo.png" },
        { "Signature", "path/to/signature.jpg" }
    }
};

// Add inventory items table
var inventoryItems = new List<Dictionary<string, string>>
{
    new Dictionary<string, string>
    {
        { "itemName", "Laptop" },
        { "quantity", "10" },
        { "price", "1200.00" }
    },
    new Dictionary<string, string>
    {
        { "itemName", "Monitor" },
        { "quantity", "15" },
        { "price", "300.00" }
    },
    new Dictionary<string, string>
    {
        { "itemName", "Keyboard" },
        { "quantity", "20" },
        { "price", "50.00" }
    }
};

data.AddTable("inventory", inventoryItems);

var template = new WordTemplate("inventory_template.docx", data);
byte[] document = template.GetFile();
```

In your Word template, you would have:

```text
Title: {{Title}}
Date: {{Date}}
Prepared By: {{PreparedBy}}

Company Logo: {% CompanyLogo %}

Inventory Items:
{{#docTable inventory}}
{{item.itemName}} {{item.quantity}} {{item.price}}
{{/docTable}}

Signature: {% Signature %}
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.
