using System.Collections.Generic;

namespace ForgeDoc.Structs;

public struct WordTemplateData
    {
        // Dictionary for simple key-value placeholders
        public Dictionary<string, string> Placeholders { get; set; }
        
        // Nested dictionary for tables with their own key-value pairs
        public Dictionary<string, Dictionary<string, string>> Tables { get; set; }
        
        // Constructor to initialize both dictionaries
        public WordTemplateData(Dictionary<string, string> placeholders, Dictionary<string, Dictionary<string, string>> tables)
        {
            Placeholders = placeholders ?? new Dictionary<string, string>();
            Tables = tables ?? new Dictionary<string, Dictionary<string, string>>();
        }
        
        // Default constructor
        public WordTemplateData()
        {
            Placeholders = new Dictionary<string, string>();
            Tables = new Dictionary<string, Dictionary<string, string>>();
        }
        
        // Method to add a placeholder
        public void AddPlaceholder(string key, string value)
        {
            Placeholders[key] = value;
        }
        
        // Method to add a table
        public void AddTable(string tableName, Dictionary<string, string> tableData)
        {
            Tables[tableName] = tableData;
        }
        
        // Method to add a key-value pair to an existing table
        public void AddTableEntry(string tableName, string key, string value)
        {
            if (!Tables.ContainsKey(tableName))
            {
                Tables[tableName] = new Dictionary<string, string>();
            }
            
            Tables[tableName][key] = value;
        }
        
        // Method to check if a placeholder exists
        public bool HasPlaceholder(string key)
        {
            return Placeholders.ContainsKey(key);
        }
        
        // Method to check if a table exists
        public bool HasTable(string tableName)
        {
            return Tables.ContainsKey(tableName);
        }
        
        // Method to check if a table entry exists
        public bool HasTableEntry(string tableName, string key)
        {
            return Tables.ContainsKey(tableName) && Tables[tableName].ContainsKey(key);
        }
    }