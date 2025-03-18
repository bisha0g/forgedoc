# ForgeDoc - Word Template Processor

ForgeDoc is a library for processing Word templates with placeholders and generating dynamic Word documents.

## Features

- Replace text placeholders in Word documents
- Add images to Word documents
- Render tables in Word documents

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
