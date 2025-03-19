using System.Collections.Generic;

namespace ForgeDoc.Structs;

public struct WordTemplateData
    {
        // Dictionary for simple key-value placeholders
        public Dictionary<string, string> Placeholders { get; set; }
        public Dictionary<string, string> Images { get; set; }
        // Dictionary for table data, where the key is the table name and the value is a list of dictionaries
        // Each dictionary in the list represents a row, with keys as column names and values as cell contents
        public Dictionary<string, List<Dictionary<string, string>>> Tables { get; set; }
        
        // Constructor to initialize all dictionaries
        public WordTemplateData(Dictionary<string, string> placeholders, Dictionary<string, string> images, Dictionary<string, List<Dictionary<string, string>>> tables = null)
        {
            Placeholders = placeholders ?? new Dictionary<string, string>();
            Tables = tables ?? new Dictionary<string, List<Dictionary<string, string>>>();
            Images = images ?? new Dictionary<string, string>();
        }
        
        // Default constructor
        public WordTemplateData()
        {
            Placeholders = new Dictionary<string, string>();
            Tables = new Dictionary<string, List<Dictionary<string, string>>>();
            Images = new Dictionary<string, string>();
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
    }