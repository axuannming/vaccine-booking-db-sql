CREATE TABLE Citizens (
    CitizenID AUTOINCREMENT PRIMARY KEY,
    HKID TEXT(20) NOT NULL,
    CitizenName TEXT(100) NOT NULL,
    FirstName TEXT(50),
    LastName TEXT(50),
    Gender TEXT(10) NOT NULL,
    DOB DATETIME NOT NULL,
    Phone TEXT(20),
    Email TEXT(100),
    [Address] TEXT(255),
    District TEXT(50),
    AllergyHistory TEXT(255),
    Occupation TEXT(100),
    PlaceOfBirth TEXT(100)
);
CREATE TABLE Vaccines (
    VaccineID AUTOINCREMENT PRIMARY KEY,
    VaccineName TEXT(100) NOT NULL,
    Manufacturer TEXT(100) NOT NULL,
    RequiredDoses LONG NOT NULL,
    DoseIntervalDays LONG NOT NULL,
    StorageTemp TEXT(50),
    MinAge LONG NOT NULL,
    MaxAge LONG NOT NULL
);
CREATE TABLE VaccinationCenters (
    CenterID AUTOINCREMENT PRIMARY KEY,
    CenterName TEXT(100) NOT NULL,
    [Address] TEXT(255) NOT NULL,
    Region TEXT(50) NOT NULL,
    ContactPhone TEXT(20),
    OperatingHours TEXT(50),
    MaxDailyService LONG NOT NULL,
    DoctorName TEXT(100),
    IsActive YESNO NOT NULL
);
CREATE TABLE VaccineBatches (
    BatchID AUTOINCREMENT PRIMARY KEY,
    BatchNumber TEXT(50) NOT NULL,
    VaccineID LONG NOT NULL,
    ManufactureDate DATETIME NOT NULL,
    ExpiryDate DATETIME NOT NULL,
    CurrentStock LONG NOT NULL,
    Supplier TEXT(100) NOT NULL
);
CREATE TABLE Users (
    UserID AUTOINCREMENT PRIMARY KEY,
    Username TEXT(50) NOT NULL,
    UserPassword TEXT(50) NOT NULL,
    [Role] TEXT(20) NOT NULL,
    LinkedCitizenID LONG,
    IsActive YESNO NOT NULL
);
CREATE TABLE AppointmentSlots (
    SlotID AUTOINCREMENT PRIMARY KEY,
    CenterID LONG NOT NULL,
    VaccineID LONG NOT NULL,
    SlotDate DATETIME NOT NULL,
    SlotTime DATETIME NOT NULL,
    Capacity LONG NOT NULL,
    BookedCount LONG NOT NULL,
    SlotStatus TEXT(20) NOT NULL
);
CREATE TABLE Appointments (
    AppointmentID AUTOINCREMENT PRIMARY KEY,
    CitizenID LONG NOT NULL,
    VaccineID LONG NOT NULL,
    CenterID LONG NOT NULL,
    SlotID LONG,
    BatchID LONG,
    AppointmentDate DATETIME NOT NULL,
    DoseNo LONG NOT NULL,
    AppointmentStatus TEXT(20) NOT NULL,
    ActualVaccinationDate DATETIME,
    VaccinatedBy LONG,
    Notes TEXT(255),
    CancellationReason TEXT(255)
);
CREATE TABLE VaccinationRecords (
    RecordID AUTOINCREMENT PRIMARY KEY,
    AppointmentID LONG NOT NULL,
    CitizenID LONG NOT NULL,
    VaccineID LONG NOT NULL,
    BatchID LONG NOT NULL,
    DoseNo LONG NOT NULL,
    VaccinationDate DATETIME NOT NULL,
    StaffUserID LONG NOT NULL,
    OutcomeStatus TEXT(30) NOT NULL
);
CREATE TABLE PreVaccinationScreening (
    ScreeningID AUTOINCREMENT PRIMARY KEY,
    AppointmentID LONG NOT NULL,
    CitizenID LONG NOT NULL,
    ScreeningDate DATETIME NOT NULL,
    RiskFlag YESNO NOT NULL,
    Temperature DOUBLE NOT NULL,
    AllergyConfirmed YESNO NOT NULL,
    ContraindicationNote TEXT(255),
    ScreenedBy LONG NOT NULL
);
CREATE TABLE ConsentRecords (
    ConsentID AUTOINCREMENT PRIMARY KEY,
    AppointmentID LONG NOT NULL,
    CitizenID LONG NOT NULL,
    ConsentStatus TEXT(20) NOT NULL,
    ConsentDate DATETIME NOT NULL,
    WitnessedBy LONG,
    ConsentNote TEXT(255)
);
CREATE TABLE ObservationRecords (
    ObservationID AUTOINCREMENT PRIMARY KEY,
    RecordID LONG NOT NULL,
    CitizenID LONG NOT NULL,
    ObservationStart DATETIME NOT NULL,
    ObservationEnd DATETIME NOT NULL,
    ImmediateReaction YESNO NOT NULL,
    ObservationOutcome TEXT(30) NOT NULL,
    ObservedBy LONG NOT NULL
);
CREATE TABLE SideEffects (
    ReportID AUTOINCREMENT PRIMARY KEY,
    VaccinationRecordID LONG,
    AppointmentID LONG,
    CitizenID LONG NOT NULL,
    SymptomType TEXT(100) NOT NULL,
    SeverityLevel TEXT(20) NOT NULL,
    ReportDate DATETIME NOT NULL,
    Details TEXT(255),
    FollowUpStatus TEXT(50)
);
CREATE TABLE AdverseEffectFollowUps (
    FollowUpID AUTOINCREMENT PRIMARY KEY,
    ReportID LONG NOT NULL,
    FollowUpDate DATETIME NOT NULL,
    StaffUserID LONG NOT NULL,
    ActionTaken TEXT(255),
    FollowUpOutcome TEXT(50),
    EscalationFlag YESNO NOT NULL
);
CREATE TABLE Notifications (
    NotificationID AUTOINCREMENT PRIMARY KEY,
    CitizenID LONG NOT NULL,
    NotificationType TEXT(30) NOT NULL,
    Message TEXT(255) NOT NULL,
    SendDate DATETIME NOT NULL,
    IsRead YESNO NOT NULL,
    RelatedAppointmentID LONG
);
CREATE TABLE VaccineAllocation (
    AllocationID AUTOINCREMENT PRIMARY KEY,
    BatchID LONG NOT NULL,
    CenterID LONG NOT NULL,
    QuantityAllocated LONG NOT NULL,
    AllocationDate DATETIME NOT NULL,
    AllocatedBy LONG NOT NULL
);
CREATE TABLE InventoryTransactions (
    TransactionID AUTOINCREMENT PRIMARY KEY,
    BatchID LONG NOT NULL,
    CenterID LONG NOT NULL,
    TransactionType TEXT(30) NOT NULL,
    Quantity LONG NOT NULL,
    TransactionDate DATETIME NOT NULL,
    RelatedRecordID LONG,
    ProcessedBy LONG NOT NULL
);
CREATE TABLE AuditLog (
    LogID AUTOINCREMENT PRIMARY KEY,
    UserID LONG NOT NULL,
    [Action] TEXT(20) NOT NULL,
    TableName TEXT(50) NOT NULL,
    RecordID LONG NOT NULL,
    ActionDate DATETIME NOT NULL,
    Details TEXT(255)
);

