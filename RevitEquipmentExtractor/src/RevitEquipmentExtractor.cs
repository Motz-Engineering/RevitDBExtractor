using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace RevitEquipmentExtractor
{
    public class RevitEquipmentExtractor
    {
        // Replace with your database connection string
        private const string ConnectionString = "Server=YourServer;Database=YourDatabase;Trusted_Connection=True;";
        
        // Main method that runs the extraction process
        public void ExtractEquipmentData()
        {
            try
            {
                // Step 1: Get project file data from database
                DataTable projectFiles = GetProjectFilesFromDatabase();
                
                // Step 2: Process each Revit file
                foreach (DataRow row in projectFiles.Rows)
                {
                    string projectNumber = row["ProjectNumber"].ToString();
                    string projectName = row["ProjectName"].ToString();
                    string revitFilePath = row["RevitFilePath"].ToString();
                    string revitVersion = row["RevitVersion"].ToString();
                    string revitFileGUID = row["RevitFileGUID"].ToString();
                    
                    Console.WriteLine($"Processing project {projectNumber}: {projectName}");
                    Console.WriteLine($"Revit file: {revitFilePath} (Version: {revitVersion})");
                    
                    // Update processing status
                    UpdateProjectFileStatus(projectNumber, "Processing");
                    
                    try
                    {
                        // Step 3: Extract equipment data from the Revit file
                        List<EquipmentData> equipmentList = ExtractEquipmentFromRevitFile(revitFilePath, revitVersion);
                        
                        // Step 4: Store the extracted data back to the database
                        StoreEquipmentDataInDatabase(projectNumber, revitFileGUID, equipmentList);
                        
                        // Update processing status to completed
                        UpdateProjectFileStatus(projectNumber, "Completed");
                        
                        Console.WriteLine($"Completed extraction for project {projectNumber}. {equipmentList.Count} equipment items found.");
                    }
                    catch (Exception ex)
                    {
                        // Update processing status to failed
                        UpdateProjectFileStatus(projectNumber, "Failed");
                        Console.WriteLine($"Error processing project {projectNumber}: {ex.Message}");
                    }
                }
                
                Console.WriteLine("Equipment data extraction complete for all projects.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in extraction process: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        // Get project files from the database
        private DataTable GetProjectFilesFromDatabase()
        {
            DataTable projectFiles = new DataTable();
            
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                
                string query = @"
                    SELECT 
                        ProjectNumber, 
                        ProjectName, 
                        RevitFilePath, 
                        RevitVersion,
                        RevitFileGUID
                    FROM 
                        ProjectFiles 
                    WHERE 
                        RevitFilePath IS NOT NULL
                        AND ProcessingStatus = 'Pending'";
                
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        adapter.Fill(projectFiles);
                    }
                }
            }
            
            Console.WriteLine($"Retrieved {projectFiles.Rows.Count} project files from database.");
            return projectFiles;
        }
        
        // Update project file processing status
        private void UpdateProjectFileStatus(string projectNumber, string status)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                
                string query = @"
                    UPDATE ProjectFiles 
                    SET 
                        ProcessingStatus = @Status,
                        LastProcessedDate = GETDATE()
                    WHERE 
                        ProjectNumber = @ProjectNumber";
                
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                    command.Parameters.AddWithValue("@Status", status);
                    command.ExecuteNonQuery();
                }
            }
        }
        
        // Extract equipment data from a Revit file
        private List<EquipmentData> ExtractEquipmentFromRevitFile(string filePath, string revitVersion)
        {
            List<EquipmentData> equipmentList = new List<EquipmentData>();
            
            try
            {
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"File not found: {filePath}");
                    return equipmentList;
                }
                
                // Initialize Revit application
                Application revitApp = new Application();
                
                // Open document with option set to audit if needed
                OpenOptions openOptions = new OpenOptions();
                openOptions.Audit = true;
                
                // Open the document
                Document document = revitApp.OpenDocumentFile(filePath, openOptions);
                
                try
                {
                    // Get all mechanical equipment elements
                    FilteredElementCollector mechanicalCollector = new FilteredElementCollector(document)
                        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                        .WhereElementIsNotElementType();
                        
                    FilteredElementCollector electricalCollector = new FilteredElementCollector(document)
                        .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                        .WhereElementIsNotElementType();
                    
                    // Process mechanical equipment
                    foreach (Element element in mechanicalCollector)
                    {
                        EquipmentData equipment = ProcessEquipmentElement(element, "Mechanical");
                        if (equipment != null)
                        {
                            equipmentList.Add(equipment);
                        }
                    }
                    
                    // Process electrical equipment
                    foreach (Element element in electricalCollector)
                    {
                        EquipmentData equipment = ProcessEquipmentElement(element, "Electrical");
                        if (equipment != null)
                        {
                            equipmentList.Add(equipment);
                        }
                    }
                }
                finally
                {
                    // Close the document
                    document.Close(false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error extracting equipment from file {filePath}: {ex.Message}");
            }
            
            return equipmentList;
        }
        
        // Process a single equipment element
        private EquipmentData ProcessEquipmentElement(Element element, string discipline)
        {
            try
            {
                // Get element ID and Family/Type names
                ElementId elementId = element.Id;
                string elementName = element.Name;
                
                // Get Equipment Designation (shared parameter)
                string equipmentDesignation = GetParameterValueAsString(element, "Equipment Designation");
                
                // Get Equipment Type (parameter)
                string equipmentType = GetParameterValueAsString(element, "Equipment Type");
                if (string.IsNullOrEmpty(equipmentType))
                {
                    // Try some alternate parameter names that might exist
                    equipmentType = GetParameterValueAsString(element, "Type");
                    if (string.IsNullOrEmpty(equipmentType))
                    {
                        equipmentType = GetParameterValueAsString(element, "Family");
                    }
                }
                
                // Check if we have the required data
                if (string.IsNullOrEmpty(equipmentDesignation) && string.IsNullOrEmpty(equipmentType))
                {
                    return null; // Skip equipment with no useful data
                }
                
                // Create equipment data
                var equipment = new EquipmentData
                {
                    ElementId = elementId.IntegerValue,
                    ElementName = elementName,
                    Discipline = discipline,
                    EquipmentDesignation = equipmentDesignation,
                    EquipmentType = equipmentType
                };
                
                // Calculate data hash
                equipment.ElementDataHash = CalculateElementDataHash(equipment);
                
                return equipment;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing element {element.Id}: {ex.Message}");
                return null;
            }
        }
        
        // Calculate hash of element data
        private string CalculateElementDataHash(EquipmentData equipment)
        {
            // Create a string containing all relevant data
            string dataToHash = $"{equipment.ElementId}|{equipment.ElementName}|{equipment.Discipline}|{equipment.EquipmentDesignation}|{equipment.EquipmentType}";
            
            // Calculate SHA-256 hash
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(dataToHash));
                return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
            }
        }
        
        // Helper method to get parameter value as string
        private string GetParameterValueAsString(Element element, string parameterName)
        {
            // First check if it's a shared parameter
            foreach (Parameter param in element.Parameters)
            {
                if (param.Definition.Name.Equals(parameterName, StringComparison.OrdinalIgnoreCase))
                {
                    return GetParameterValueAsString(param);
                }
            }
            
            // Then check instance parameters
            Parameter parameter = element.LookupParameter(parameterName);
            if (parameter != null && parameter.HasValue)
            {
                return GetParameterValueAsString(parameter);
            }
            
            // Check element type parameters as a fallback
            ElementId typeId = element.GetTypeId();
            if (typeId != ElementId.InvalidElementId)
            {
                Element elementType = element.Document.GetElement(typeId);
                if (elementType != null)
                {
                    parameter = elementType.LookupParameter(parameterName);
                    if (parameter != null && parameter.HasValue)
                    {
                        return GetParameterValueAsString(parameter);
                    }
                }
            }
            
            return string.Empty;
        }
        
        // Helper method to extract parameter value based on its storage type
        private string GetParameterValueAsString(Parameter parameter)
        {
            if (parameter == null || !parameter.HasValue)
            {
                return string.Empty;
            }
            
            switch (parameter.StorageType)
            {
                case StorageType.String:
                    return parameter.AsString();
                case StorageType.Integer:
                    return parameter.AsInteger().ToString();
                case StorageType.Double:
                    return parameter.AsDouble().ToString();
                case StorageType.ElementId:
                    return parameter.AsElementId().IntegerValue.ToString();
                default:
                    return string.Empty;
            }
        }
        
        // Store equipment data in database
        private void StoreEquipmentDataInDatabase(string projectNumber, string revitFileGUID, List<EquipmentData> equipmentList)
        {
            using (SqlConnection connection = new SqlConnection(ConnectionString))
            {
                connection.Open();
                
                foreach (EquipmentData equipment in equipmentList)
                {
                    using (SqlCommand command = new SqlCommand("sp_UpdateEquipmentRecord", connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        
                        command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                        command.Parameters.AddWithValue("@RevitFileGUID", revitFileGUID);
                        command.Parameters.AddWithValue("@ElementId", equipment.ElementId);
                        command.Parameters.AddWithValue("@ElementName", equipment.ElementName ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@Discipline", equipment.Discipline);
                        command.Parameters.AddWithValue("@EquipmentDesignation", equipment.EquipmentDesignation ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@EquipmentType", equipment.EquipmentType ?? (object)DBNull.Value);
                        command.Parameters.AddWithValue("@ElementDataHash", equipment.ElementDataHash);
                        
                        command.ExecuteNonQuery();
                    }
                }
            }
        }
    }
    
    // Class to hold equipment data
    public class EquipmentData
    {
        public int ElementId { get; set; }
        public string ElementName { get; set; }
        public string Discipline { get; set; } // Mechanical or Electrical
        public string EquipmentDesignation { get; set; }
        public string EquipmentType { get; set; }
        public string ElementDataHash { get; set; }
    }
    
    // Program entry point
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Revit Equipment Extractor");
            Console.WriteLine("-------------------------");
            
            RevitEquipmentExtractor extractor = new RevitEquipmentExtractor();
            extractor.ExtractEquipmentData();
            
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
} 