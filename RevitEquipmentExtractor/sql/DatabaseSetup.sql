-- Access Database Setup Script
-- This script should be run in Access to create the necessary tables and relationships

-- Note: FilepathTable already exists with columns:
-- ID, ProjectNum, Filepath, RevitVersion

-- Create ProjectEquipment table with change tracking
CREATE TABLE ProjectEquipment (
    EquipmentId AUTOINCREMENT PRIMARY KEY,
    ProjectNumber TEXT(50),
    RevitFileGUID TEXT(50),
    ElementId LONG,
    Equipment_Designation TEXT(100),
    ElementCategory TEXT(50),
    ElementDataHash TEXT(64),
    RecordStatus TEXT(20),
    VersionNumber LONG,
    FirstExtractedDate DATETIME,
    LastExtractedDate DATETIME,
    LastUpdatedDate DATETIME
);

-- Create indexes for better query performance
CREATE INDEX IX_ProjectEquipment_ProjectNumber ON ProjectEquipment(ProjectNumber);
CREATE INDEX IX_ProjectEquipment_ElementId ON ProjectEquipment(ElementId);
CREATE INDEX IX_ProjectEquipment_Equipment_Designation ON ProjectEquipment(Equipment_Designation);
CREATE INDEX IX_ProjectEquipment_ElementCategory ON ProjectEquipment(ElementCategory);
CREATE INDEX IX_ProjectEquipment_RecordStatus ON ProjectEquipment(RecordStatus);
CREATE INDEX IX_ProjectEquipment_RevitFileGUID ON ProjectEquipment(RevitFileGUID);

-- Create relationship between tables
ALTER TABLE ProjectEquipment
ADD CONSTRAINT FK_ProjectEquipment_FilepathTable
FOREIGN KEY (ProjectNumber) REFERENCES FilepathTable(ProjectNum);

-- Create a view for active equipment records
CREATE VIEW vw_ActiveEquipment AS
SELECT 
    EquipmentId,
    ProjectNumber,
    RevitFileGUID,
    ElementId,
    Equipment_Designation,
    ElementCategory,
    ElementDataHash,
    VersionNumber,
    FirstExtractedDate,
    LastExtractedDate,
    LastUpdatedDate
FROM 
    ProjectEquipment
WHERE 
    RecordStatus = 'Active';

-- Create a stored procedure for updating equipment records
CREATE PROCEDURE sp_UpdateEquipmentRecord
    @ProjectNumber VARCHAR(50),
    @RevitFileGUID VARCHAR(50),
    @ElementId INT,
    @Equipment_Designation VARCHAR(100),
    @ElementCategory VARCHAR(50),
    @ElementDataHash VARCHAR(64)
AS
BEGIN
    SET NOCOUNT ON;
    
    DECLARE @CurrentVersion INT;
    DECLARE @CurrentHash VARCHAR(64);
    
    -- Get current version and hash
    SELECT TOP 1 
        @CurrentVersion = VersionNumber,
        @CurrentHash = ElementDataHash
    FROM 
        ProjectEquipment
    WHERE 
        ProjectNumber = @ProjectNumber 
        AND Equipment_Designation = @Equipment_Designation
        AND RecordStatus = 'Active'
    ORDER BY 
        VersionNumber DESC;
    
    -- If no record exists, insert new
    IF @CurrentVersion IS NULL
    BEGIN
        INSERT INTO ProjectEquipment (
            ProjectNumber,
            RevitFileGUID,
            ElementId,
            Equipment_Designation,
            ElementCategory,
            ElementDataHash
        )
        VALUES (
            @ProjectNumber,
            @RevitFileGUID,
            @ElementId,
            @Equipment_Designation,
            @ElementCategory,
            @ElementDataHash
        );
    END
    -- If hash is different, mark current as modified and insert new version
    ELSE IF @CurrentHash <> @ElementDataHash
    BEGIN
        -- Update current record status
        UPDATE ProjectEquipment
        SET 
            RecordStatus = 'Modified',
            LastUpdatedDate = GETDATE()
        WHERE 
            ProjectNumber = @ProjectNumber 
            AND Equipment_Designation = @Equipment_Designation
            AND VersionNumber = @CurrentVersion;
        
        -- Insert new version
        INSERT INTO ProjectEquipment (
            ProjectNumber,
            RevitFileGUID,
            ElementId,
            Equipment_Designation,
            ElementCategory,
            ElementDataHash,
            VersionNumber
        )
        VALUES (
            @ProjectNumber,
            @RevitFileGUID,
            @ElementId,
            @Equipment_Designation,
            @ElementCategory,
            @ElementDataHash,
            @CurrentVersion + 1
        );
    END
    -- If hash is same, just update last extracted date
    ELSE
    BEGIN
        UPDATE ProjectEquipment
        SET 
            LastExtractedDate = GETDATE()
        WHERE 
            ProjectNumber = @ProjectNumber 
            AND Equipment_Designation = @Equipment_Designation
            AND VersionNumber = @CurrentVersion;
    END
END; 