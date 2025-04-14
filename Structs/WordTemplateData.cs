using System.Collections.Generic;

namespace ForgeDoc.Structs;

public struct WordTemplateData
    {
        // Dictionary for simple key-value placeholders
        public Dictionary<string, string> Placeholders { get; set; }
        public Dictionary<string, string> Images { get; set; }
        // Dictionary for header images
        public Dictionary<string, string> HeaderImages { get; set; }
        // Dictionary for table data, where the key is the table name and the value is a list of dictionaries
        // Each dictionary in the list represents a row, with keys as column names and values as cell contents
        public Dictionary<string, List<Dictionary<string, string>>> Tables { get; set; }
        
        // Dictionary for special character placeholders with specific fonts
        public Dictionary<string, (string Character, string Font)> SpecialCharacters { get; set; }
        
        // Constructor to initialize all dictionaries
        public WordTemplateData(Dictionary<string, string> placeholders, Dictionary<string, string> images, Dictionary<string, List<Dictionary<string, string>>> tables = null)
        {
            Placeholders = placeholders ?? new Dictionary<string, string>();
            Tables = tables ?? new Dictionary<string, List<Dictionary<string, string>>>();
            Images = images ?? new Dictionary<string, string>();
            HeaderImages = new Dictionary<string, string>();
            SpecialCharacters = new Dictionary<string, (string Character, string Font)>();
        }
        
        // Default constructor
        public WordTemplateData()
        {
            Placeholders = new Dictionary<string, string>();
            Tables = new Dictionary<string, List<Dictionary<string, string>>>();
            Images = new Dictionary<string, string>();
            HeaderImages = new Dictionary<string, string>();
            SpecialCharacters = new Dictionary<string, (string Character, string Font)>();
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
        
        // Method to add a header image
        public void AddHeaderImage(string key, string imagePath)
        {
            HeaderImages[key] = imagePath;
        }
        
        // Method to check if a header image exists
        public bool HasHeaderImage(string key)
        {
            return HeaderImages.ContainsKey(key);
        }
        
        // Method to get a header image path
        public string GetHeaderImage(string key)
        {
            return HeaderImages.ContainsKey(key) ? HeaderImages[key] : null;
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
    }