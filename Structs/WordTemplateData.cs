using System.Collections.Generic;

namespace ForgeDoc.Structs;

public struct WordTemplateData
    {
        // Dictionary for simple key-value placeholders
        public Dictionary<string, string> Placeholders { get; set; }
        
        // Constructor to initialize both dictionaries
        public WordTemplateData(Dictionary<string, string> placeholders, Dictionary<string, Dictionary<string, string>> tables)
        {
            Placeholders = placeholders ?? new Dictionary<string, string>();
        }
        
        // Default constructor
        public WordTemplateData()
        {
            Placeholders = new Dictionary<string, string>();
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
    }