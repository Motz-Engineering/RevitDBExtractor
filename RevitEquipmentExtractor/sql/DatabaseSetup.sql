-- Create ProjectFiles table
CREATE TABLE ProjectFiles (
    ProjectNumber VARCHAR(50) PRIMARY KEY,
    ProjectName VARCHAR(100) NOT NULL,
    RevitFilePath NVARCHAR(500),
    RevitVersion VARCHAR(50),
    RevitFileGUID VARCHAR(50) NOT NULL,  -- Unique identifier for the file
    LastProcessedDate DATETIME NULL,
    ProcessingStatus VARCHAR(20) DEFAULT 'Pending'  -- Pending, Processing, Completed, Failed
);

-- Create ProjectEquipment table with change tracking
CREATE TABLE ProjectEquipment (
    EquipmentId INT IDENTITY(1,1) NOT NULL PRIMARY KEY,
    ProjectNumber VARCHAR(50) NOT NULL,
    RevitFileGUID VARCHAR(50) NOT NULL,  -- Unique identifier for the file
    ElementId INT NOT NULL,
    ElementName VARCHAR(100) NULL,
    Discipline VARCHAR(50) NOT NULL,
    EquipmentDesignation VARCHAR(100) NULL,
    EquipmentType VARCHAR(100) NULL,
    ElementDataHash VARCHAR(64) NOT NULL,  -- Hash of parameter values
    RecordStatus VARCHAR(20) NOT NULL DEFAULT 'Active',  -- Active, Modified, Deleted
    VersionNumber INT NOT NULL DEFAULT 1,
    FirstExtractedDate DATETIME NOT NULL DEFAULT GETDATE(),
    LastExtractedDate DATETIME NOT NULL DEFAULT GETDATE(),
    LastUpdatedDate DATETIME NULL,
    CONSTRAINT UK_ProjectEquipment_Element UNIQUE (ProjectNumber, ElementId, VersionNumber),
    FOREIGN KEY (ProjectNumber) REFERENCES ProjectFiles(ProjectNumber)
);

-- Create indexes for better query performance
CREATE INDEX IX_ProjectEquipment_ProjectNumber ON ProjectEquipment(ProjectNumber);
CREATE INDEX IX_ProjectEquipment_ElementId ON ProjectEquipment(ElementId);
CREATE INDEX IX_ProjectEquipment_Discipline ON ProjectEquipment(Discipline);
CREATE INDEX IX_ProjectEquipment_EquipmentType ON ProjectEquipment(EquipmentType);
CREATE INDEX IX_ProjectEquipment_RecordStatus ON ProjectEquipment(RecordStatus);
CREATE INDEX IX_ProjectEquipment_RevitFileGUID ON ProjectEquipment(RevitFileGUID);

-- Create a view for active equipment records
CREATE VIEW vw_ActiveEquipment AS
SELECT 
    EquipmentId,
    ProjectNumber,
    RevitFileGUID,
    ElementId,
    ElementName,
    Discipline,
    EquipmentDesignation,
    EquipmentType,
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
    @ElementName VARCHAR(100),
    @Discipline VARCHAR(50),
    @EquipmentDesignation VARCHAR(100),
    @EquipmentType VARCHAR(100),
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
        AND ElementId = @ElementId
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
            ElementName,
            Discipline,
            EquipmentDesignation,
            EquipmentType,
            ElementDataHash
        )
        VALUES (
            @ProjectNumber,
            @RevitFileGUID,
            @ElementId,
            @ElementName,
            @Discipline,
            @EquipmentDesignation,
            @EquipmentType,
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
            AND ElementId = @ElementId
            AND VersionNumber = @CurrentVersion;
        
        -- Insert new version
        INSERT INTO ProjectEquipment (
            ProjectNumber,
            RevitFileGUID,
            ElementId,
            ElementName,
            Discipline,
            EquipmentDesignation,
            EquipmentType,
            ElementDataHash,
            VersionNumber
        )
        VALUES (
            @ProjectNumber,
            @RevitFileGUID,
            @ElementId,
            @ElementName,
            @Discipline,
            @EquipmentDesignation,
            @EquipmentType,
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
            AND ElementId = @ElementId
            AND VersionNumber = @CurrentVersion;
    END
END; 