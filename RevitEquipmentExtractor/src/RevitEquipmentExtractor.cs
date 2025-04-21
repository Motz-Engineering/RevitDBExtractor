using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
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
        // Access database connection string
        private const string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=S:\Engineering\_00 Transmittal and Deliverable Logs\Deliverables-Testing - Copy.accdb;";
        
        // Dictionary to map Revit versions to their corresponding API versions
        private static readonly Dictionary<string, string> RevitVersionMap = new Dictionary<string, string>
        {
            { "2021", "Revit2021" },
            { "2022", "Revit2022" },
            { "2023", "Revit2023" },
            { "2024", "Revit2024" },
            { "2025", "Revit2025" }
        };
        
        // Main method that runs the extraction process
        public void ExtractEquipmentData(string specificProjectNumber = null)
        {
            try
            {
                // Step 1: Get project file data from database
                DataTable projectFiles = GetProjectFilesFromDatabase(specificProjectNumber);
                
                if (projectFiles.Rows.Count == 0)
                {
                    if (specificProjectNumber != null)
                    {
                        Console.WriteLine($"No project found with number: {specificProjectNumber}");
                    }
                    else
                    {
                        Console.WriteLine("No projects found in database.");
                    }
                    return;
                }
                
                // Step 2: Process each Revit file
                foreach (DataRow row in projectFiles.Rows)
                {
                    string projectNumber = row["ProjectNum"].ToString();
                    string basePath = row["Filepath"].ToString();
                    string revitVersion = row["RevitVersion"].ToString();
                    
                    Console.WriteLine($"Processing project {projectNumber}");
                    Console.WriteLine($"Base path: {basePath}");
                    Console.WriteLine($"Revit version: {revitVersion}");
                    
                    try
                    {
                        // Verify the Revit version is supported
                        if (!RevitVersionMap.ContainsKey(revitVersion))
                        {
                            Console.WriteLine($"Unsupported Revit version {revitVersion} for project {projectNumber}. Skipping.");
                            continue;
                        }
                        
                        // Step 3: Find and process Revit files
                        List<string> revitFiles = FindRevitFiles(basePath);
                        
                        if (revitFiles.Count == 0)
                        {
                            Console.WriteLine($"No Revit files found for project {projectNumber}. Continuing to next project.");
                            continue;
                        }
                        
                        foreach (string revitFilePath in revitFiles)
                        {
                            try
                            {
                                // Generate a unique GUID for this file
                                string revitFileGUID = Guid.NewGuid().ToString();
                                
                                // Step 4: Extract equipment data from the Revit file
                                List<EquipmentData> equipmentList = ExtractEquipmentFromRevitFile(revitFilePath, revitVersion);
                                
                                // Step 5: Store the extracted data back to the database
                                StoreEquipmentDataInDatabase(projectNumber, revitFileGUID, equipmentList);
                                
                                Console.WriteLine($"Completed extraction for file {Path.GetFileName(revitFilePath)}. {equipmentList.Count} equipment items found.");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing file {revitFilePath}: {ex.Message}");
                                // Continue to next file
                                continue;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing project {projectNumber}: {ex.Message}");
                        // Continue to next project
                        continue;
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
        private DataTable GetProjectFilesFromDatabase(string specificProjectNumber = null)
        {
            DataTable projectFiles = new DataTable();
            
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                
                string query = @"
                    SELECT 
                        ProjectNum, 
                        Filepath, 
                        RevitVersion
                    FROM 
                        FilepathTable
                    WHERE 
                        Filepath IS NOT NULL 
                        AND Filepath <> ''
                        AND RevitVersion IS NOT NULL
                        AND RevitVersion <> 'N/A'
                        AND RevitVersion <> ''
                        AND CInt(RevitVersion) >= 2021
                        AND CInt(RevitVersion) <= 2025";
                
                // Add project number filter if specified
                if (!string.IsNullOrEmpty(specificProjectNumber))
                {
                    query += " AND ProjectNum = @ProjectNum";
                }
                
                using (OleDbCommand command = new OleDbCommand(query, connection))
                {
                    if (!string.IsNullOrEmpty(specificProjectNumber))
                    {
                        command.Parameters.AddWithValue("@ProjectNum", specificProjectNumber);
                    }
                    
                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        adapter.Fill(projectFiles);
                    }
                }
            }
            
            Console.WriteLine($"Retrieved {projectFiles.Rows.Count} project files from database.");
            return projectFiles;
        }
        
        // Find Revit files in the specified path pattern
        private List<string> FindRevitFiles(string basePath)
        {
            List<string> revitFiles = new List<string>();
            
            try
            {
                if (string.IsNullOrEmpty(basePath))
                {
                    Console.WriteLine("Base path is empty or null.");
                    return revitFiles;
                }

                // Construct the search pattern
                string searchPath = Path.Combine(basePath, "06 Revit");
                
                if (!Directory.Exists(searchPath))
                {
                    Console.WriteLine($"Directory not found: {searchPath}");
                    return revitFiles;
                }
                
                // Find all MEP folders
                string[] mepFolders = Directory.GetDirectories(searchPath, "*MEP*", SearchOption.TopDirectoryOnly);
                
                if (mepFolders.Length == 0)
                {
                    Console.WriteLine($"No MEP folders found in {searchPath}");
                    return revitFiles;
                }
                
                foreach (string mepFolder in mepFolders)
                {
                    try
                    {
                        // Find all .rvt files in the MEP folder
                        string[] files = Directory.GetFiles(mepFolder, "*.rvt", SearchOption.AllDirectories);
                        revitFiles.AddRange(files);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error searching in MEP folder {mepFolder}: {ex.Message}");
                        continue;
                    }
                }
                
                Console.WriteLine($"Found {revitFiles.Count} Revit files in {searchPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error finding Revit files: {ex.Message}");
            }
            
            return revitFiles;
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
                    // Get all equipment elements from the three categories
                    FilteredElementCollector mechanicalCollector = new FilteredElementCollector(document)
                        .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                        .WhereElementIsNotElementType();
                        
                    FilteredElementCollector plumbingCollector = new FilteredElementCollector(document)
                        .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
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
                    
                    // Process plumbing equipment
                    foreach (Element element in plumbingCollector)
                    {
                        EquipmentData equipment = ProcessEquipmentElement(element, "Plumbing");
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
        private EquipmentData ProcessEquipmentElement(Element element, string category)
        {
            try
            {
                // Get element ID
                ElementId elementId = element.Id;
                
                // Get Equipment Designation (required parameter)
                string equipmentDesignation = GetParameterValueAsString(element, "Equipment_Designation");
                
                // Skip if no equipment designation
                if (string.IsNullOrEmpty(equipmentDesignation))
                {
                    Console.WriteLine($"Warning: Element {elementId} has no Equipment_Designation. Skipping.");
                    return null;
                }
                
                // Create equipment data
                var equipment = new EquipmentData
                {
                    ElementId = elementId.IntegerValue,
                    Equipment_Designation = equipmentDesignation,
                    ElementCategory = category
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
            string dataToHash = $"{equipment.ElementId}|{equipment.Equipment_Designation}|{equipment.ElementCategory}";
            
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
            using (OleDbConnection connection = new OleDbConnection(ConnectionString))
            {
                connection.Open();
                
                foreach (EquipmentData equipment in equipmentList)
                {
                    // First, check if we need to update an existing record
                    string checkQuery = @"
                        SELECT TOP 1 ElementDataHash, VersionNumber 
                        FROM ProjectEquipment 
                        WHERE ProjectNumber = @ProjectNumber 
                        AND Equipment_Designation = @Equipment_Designation 
                        AND RecordStatus = 'Active' 
                        ORDER BY VersionNumber DESC";
                    
                    string currentHash = null;
                    int currentVersion = 0;
                    
                    using (OleDbCommand checkCommand = new OleDbCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                        checkCommand.Parameters.AddWithValue("@Equipment_Designation", equipment.Equipment_Designation);
                        
                        using (OleDbDataReader reader = checkCommand.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                currentHash = reader["ElementDataHash"].ToString();
                                currentVersion = Convert.ToInt32(reader["VersionNumber"]);
                            }
                        }
                    }
                    
                    if (currentHash == null)
                    {
                        // Insert new record
                        string insertQuery = @"
                            INSERT INTO ProjectEquipment (
                                ProjectNumber, RevitFileGUID, ElementId, 
                                Equipment_Designation, ElementCategory, ElementDataHash,
                                RecordStatus, VersionNumber, FirstExtractedDate, LastExtractedDate
                            ) VALUES (
                                @ProjectNumber, @RevitFileGUID, @ElementId,
                                @Equipment_Designation, @ElementCategory, @ElementDataHash,
                                'Active', 1, NOW(), NOW()
                            )";
                        
                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                            command.Parameters.AddWithValue("@RevitFileGUID", revitFileGUID);
                            command.Parameters.AddWithValue("@ElementId", equipment.ElementId);
                            command.Parameters.AddWithValue("@Equipment_Designation", equipment.Equipment_Designation);
                            command.Parameters.AddWithValue("@ElementCategory", equipment.ElementCategory);
                            command.Parameters.AddWithValue("@ElementDataHash", equipment.ElementDataHash);
                            command.ExecuteNonQuery();
                        }
                    }
                    else if (currentHash != equipment.ElementDataHash)
                    {
                        // Update existing record status
                        string updateQuery = @"
                            UPDATE ProjectEquipment 
                            SET RecordStatus = 'Modified',
                                LastUpdatedDate = NOW()
                            WHERE ProjectNumber = @ProjectNumber 
                            AND Equipment_Designation = @Equipment_Designation 
                            AND VersionNumber = @VersionNumber";
                        
                        using (OleDbCommand command = new OleDbCommand(updateQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                            command.Parameters.AddWithValue("@Equipment_Designation", equipment.Equipment_Designation);
                            command.Parameters.AddWithValue("@VersionNumber", currentVersion);
                            command.ExecuteNonQuery();
                        }
                        
                        // Insert new version
                        string insertQuery = @"
                            INSERT INTO ProjectEquipment (
                                ProjectNumber, RevitFileGUID, ElementId, 
                                Equipment_Designation, ElementCategory, ElementDataHash,
                                RecordStatus, VersionNumber, FirstExtractedDate, LastExtractedDate
                            ) VALUES (
                                @ProjectNumber, @RevitFileGUID, @ElementId,
                                @Equipment_Designation, @ElementCategory, @ElementDataHash,
                                'Active', @NewVersion, NOW(), NOW()
                            )";
                        
                        using (OleDbCommand command = new OleDbCommand(insertQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                            command.Parameters.AddWithValue("@RevitFileGUID", revitFileGUID);
                            command.Parameters.AddWithValue("@ElementId", equipment.ElementId);
                            command.Parameters.AddWithValue("@Equipment_Designation", equipment.Equipment_Designation);
                            command.Parameters.AddWithValue("@ElementCategory", equipment.ElementCategory);
                            command.Parameters.AddWithValue("@ElementDataHash", equipment.ElementDataHash);
                            command.Parameters.AddWithValue("@NewVersion", currentVersion + 1);
                            command.ExecuteNonQuery();
                        }
                    }
                    else
                    {
                        // Just update the last extracted date
                        string updateQuery = @"
                            UPDATE ProjectEquipment 
                            SET LastExtractedDate = NOW()
                            WHERE ProjectNumber = @ProjectNumber 
                            AND Equipment_Designation = @Equipment_Designation 
                            AND VersionNumber = @VersionNumber";
                        
                        using (OleDbCommand command = new OleDbCommand(updateQuery, connection))
                        {
                            command.Parameters.AddWithValue("@ProjectNumber", projectNumber);
                            command.Parameters.AddWithValue("@Equipment_Designation", equipment.Equipment_Designation);
                            command.Parameters.AddWithValue("@VersionNumber", currentVersion);
                            command.ExecuteNonQuery();
                        }
                    }
                }
            }
        }
    }
    
    // Class to hold equipment data
    public class EquipmentData
    {
        public int ElementId { get; set; }
        public string Equipment_Designation { get; set; }
        public string ElementCategory { get; set; } // Mechanical, Plumbing, or Electrical
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
            
            // Check if a specific project number was provided
            string projectNumber = args.Length > 0 ? args[0] : null;
            
            if (projectNumber != null)
            {
                Console.WriteLine($"Running extraction for project: {projectNumber}");
                extractor.ExtractEquipmentData(projectNumber);
            }
            else
            {
                Console.WriteLine("No project number specified. Running extraction for all projects.");
                extractor.ExtractEquipmentData();
            }
            
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }
    }
} 