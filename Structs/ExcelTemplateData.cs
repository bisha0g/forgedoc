using System.Collections.Generic;

namespace ForgeDoc.Structs;

public struct ExcelTemplateData
{
    // Dictionary for simple key-value placeholders
    public Dictionary<string, string> Placeholders { get; set; }
    
    // Dictionary for table data, where the key is the table name and the value is a list of dictionaries
    // Each dictionary in the list represents a row, with keys as column names and values as cell contents
    public Dictionary<string, List<Dictionary<string, string>>> Tables { get; set; }
    
    // Dictionary for image paths
    public Dictionary<string, string> Images { get; set; }
    
    // Dictionary for special character placeholders with specific fonts
    public Dictionary<string, (string Character, string Font)> SpecialCharacters { get; set; }
    
    // Dictionary for collections used in for loops
    public Dictionary<string, List<Dictionary<string, string>>> Collections { get; set; }
    
    // Dictionary for rich text placeholders
    public Dictionary<string, string> RichTextPlaceholders { get; set; }
    
    // Constructor to initialize all dictionaries
    public ExcelTemplateData(Dictionary<string, string> placeholders, Dictionary<string, string> images, Dictionary<string, List<Dictionary<string, string>>> tables = null)
    {
        Placeholders = placeholders ?? new Dictionary<string, string>();
        Tables = tables ?? new Dictionary<string, List<Dictionary<string, string>>>();
        Images = images ?? new Dictionary<string, string>();
        SpecialCharacters = new Dictionary<string, (string Character, string Font)>();
        Collections = new Dictionary<string, List<Dictionary<string, string>>>();
        RichTextPlaceholders = new Dictionary<string, string>();
    }
    
    // Default constructor
    public ExcelTemplateData()
    {
        Placeholders = new Dictionary<string, string>();
        Tables = new Dictionary<string, List<Dictionary<string, string>>>();
        Images = new Dictionary<string, string>();
        SpecialCharacters = new Dictionary<string, (string Character, string Font)>();
        Collections = new Dictionary<string, List<Dictionary<string, string>>>();
        RichTextPlaceholders = new Dictionary<string, string>();
    }
    
    // Method to add a placeholder
    public void AddPlaceholder(string key, string value)
    {
        Placeholders[key] = value;
    }
    
    // Method to check if a placeholder exists
    public bool HasPlaceholder(string key)
    {
        return Placeholders.ContainsKey(key);
    }
    
    // Method to add an image
    public void AddImage(string key, string imagePath)
    {
        Images[key] = imagePath;
    }
    
    // Method to check if an image exists
    public bool HasImage(string key)
    {
        return Images.ContainsKey(key);
    }
    
    // Method to get an image path
    public string GetImage(string key)
    {
        return Images.ContainsKey(key) ? Images[key] : null;
    }
    
    // Method to add a table
    public void AddTable(string tableName, List<Dictionary<string, string>> tableData)
    {
        if (Tables == null)
        {
            Tables = new Dictionary<string, List<Dictionary<string, string>>>();
        }
        
        Tables[tableName] = tableData;
    }
    
    // Method to add a row to an existing table
    public void AddTableRow(string tableName, Dictionary<string, string> rowData)
    {
        if (Tables == null)
        {
            Tables = new Dictionary<string, List<Dictionary<string, string>>>();
        }
        
        if (!Tables.ContainsKey(tableName))
        {
            Tables[tableName] = new List<Dictionary<string, string>>();
        }
        
        Tables[tableName].Add(rowData);
    }
    
    // Method to check if a table exists
    public bool HasTable(string tableName)
    {
        return Tables.ContainsKey(tableName);
    }
    
    // Method to add a special character with a specific font
    public void AddSpecialCharacter(string key, string character, string font)
    {
        SpecialCharacters[key] = (character, font);
    }
    
    // Method to check if a special character exists
    public bool HasSpecialCharacter(string key)
    {
        return SpecialCharacters.ContainsKey(key);
    }
    
    // Method to get a special character and its font
    public (string Character, string Font) GetSpecialCharacter(string key)
    {
        return SpecialCharacters.ContainsKey(key) ? SpecialCharacters[key] : (null, null);
    }
    
    // Method to add a collection for for-loop processing
    public void AddCollection(string name, List<Dictionary<string, string>> collection)
    {
        Collections[name] = collection;
    }
    
    // Method to add an item to a collection
    public void AddCollectionItem(string name, Dictionary<string, string> item)
    {
        if (!Collections.ContainsKey(name))
        {
            Collections[name] = new List<Dictionary<string, string>>();
        }
        
        Collections[name].Add(item);
    }
    
    // Method to check if a collection exists
    public bool HasCollection(string name)
    {
        return Collections.ContainsKey(name);
    }
    
    // Method to get a collection
    public List<Dictionary<string, string>> GetCollection(string name)
    {
        return Collections.ContainsKey(name) ? Collections[name] : null;
    }
    
    // Method to add rich text placeholder
    public void AddRichTextPlaceholder(string key, string htmlContent)
    {
        RichTextPlaceholders[key] = htmlContent;
    }
}