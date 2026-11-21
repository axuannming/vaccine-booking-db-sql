Add-Type -AssemblyName System.Drawing
$ErrorActionPreference = 'Stop'

$databasePath = '.\gp.accdb'
$backupPath = '.\gp_before_v2_data_upgrade.accdb'
$sqlRoot = '..\sql'
$querySqlRoot = '..\sql\queries'
$logPath = '.\build_access_v2_data.log'
$dbFailOnError = 128
$today = Get-Date '2026-03-25'

Set-Content -Path $logPath -Value @('START ' + (Get-Date).ToString('s'))

function Log([string]$message) {
    Add-Content -Path $logPath -Value ((Get-Date).ToString('s') + ' ' + $message)
}

function Escape-SqlString([string]$value) {
    if ($null -eq $value) { return '' }
    return $value.Replace("'", "''")
}

function Convert-ToAccessSqlValue($value) {
    if ($null -eq $value) { return 'NULL' }
    if ($value -is [bool]) { return $(if ($value) { 'True' } else { 'False' }) }
    if ($value -is [datetime]) { return '#' + $value.ToString('MM/dd/yyyy HH:mm') + '#' }
    if ($value -is [string]) { return "'" + (Escape-SqlString $value) + "'" }
    return [string]$value
}

function New-InsertSql([string]$tableName, [string[]]$columns, [hashtable]$row) {
    $columnList = ($columns | ForEach-Object { '[' + $_ + ']' }) -join ', '
    $valueList = ($columns | ForEach-Object { Convert-ToAccessSqlValue $row[$_] }) -join ', '
    return "INSERT INTO [$tableName] ($columnList) VALUES ($valueList);"
}

function Execute-Insert($db, [string]$tableName, [string[]]$columns, [hashtable]$row, [System.Collections.Generic.List[string]]$sqlOut) {
    $sql = New-InsertSql $tableName $columns $row
    $db.Execute($sql, $dbFailOnError)
    $sqlOut.Add($sql) | Out-Null
}

function Remove-QueryIfExists($db, [string]$queryName) {
    try { $db.QueryDefs.Delete($queryName) } catch { }
}

function Add-OrReplaceQuery($db, [string]$queryName, [string]$sql, [System.Collections.Specialized.OrderedDictionary]$querySqlMap) {
    Remove-QueryIfExists $db $queryName
    [void]$db.CreateQueryDef($queryName, $sql)
    $querySqlMap[$queryName] = $sql.Trim()
    Log ('QUERY READY ' + $queryName)
}

function Remove-TableIfExists($db, [string]$tableName) {
    try {
        $db.TableDefs.Delete($tableName)
        Log ('TABLE DROPPED ' + $tableName)
    } catch {
    }
}

function Get-Age([datetime]$dob) {
    $age = $today.Year - $dob.Year
    if ($dob.AddYears($age) -gt $today) { $age-- }
    return $age
}

function Get-AgeGroup([datetime]$dob) {
    $age = Get-Age $dob
    if ($age -le 11) { return '0-11' }
    if ($age -le 17) { return '12-17' }
    if ($age -le 29) { return '18-29' }
    if ($age -le 44) { return '30-44' }
    if ($age -le 59) { return '45-59' }
    return '60+'
}

function Get-DobFromAge([int]$age, [int]$offsetDays = 0) {
    return $today.AddYears(-$age).AddDays(-$offsetDays)
}

function Resolve-BatchId([int]$vaccineId, [int]$doseNo) {
    if ($doseNo -ge 2 -and $vaccineId -le 10) {
        return $vaccineId + 30
    }
    return $vaccineId
}

function Build-CreateSql() {
    return @(
@'
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
'@,
@'
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
'@,
@'
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
'@,
@'
CREATE TABLE VaccineBatches (
    BatchID AUTOINCREMENT PRIMARY KEY,
    BatchNumber TEXT(50) NOT NULL,
    VaccineID LONG NOT NULL,
    ManufactureDate DATETIME NOT NULL,
    ExpiryDate DATETIME NOT NULL,
    CurrentStock LONG NOT NULL,
    Supplier TEXT(100) NOT NULL
);
'@,
@'
CREATE TABLE Users (
    UserID AUTOINCREMENT PRIMARY KEY,
    Username TEXT(50) NOT NULL,
    UserPassword TEXT(50) NOT NULL,
    [Role] TEXT(20) NOT NULL,
    LinkedCitizenID LONG,
    IsActive YESNO NOT NULL
);
'@,
@'
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
'@,
@'
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
'@,
@'
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
'@,
@'
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
'@,
@'
CREATE TABLE ConsentRecords (
    ConsentID AUTOINCREMENT PRIMARY KEY,
    AppointmentID LONG NOT NULL,
    CitizenID LONG NOT NULL,
    ConsentStatus TEXT(20) NOT NULL,
    ConsentDate DATETIME NOT NULL,
    WitnessedBy LONG,
    ConsentNote TEXT(255)
);
'@,
@'
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
'@,
@'
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
'@,
@'
CREATE TABLE AdverseEffectFollowUps (
    FollowUpID AUTOINCREMENT PRIMARY KEY,
    ReportID LONG NOT NULL,
    FollowUpDate DATETIME NOT NULL,
    StaffUserID LONG NOT NULL,
    ActionTaken TEXT(255),
    FollowUpOutcome TEXT(50),
    EscalationFlag YESNO NOT NULL
);
'@,
@'
CREATE TABLE Notifications (
    NotificationID AUTOINCREMENT PRIMARY KEY,
    CitizenID LONG NOT NULL,
    NotificationType TEXT(30) NOT NULL,
    Message TEXT(255) NOT NULL,
    SendDate DATETIME NOT NULL,
    IsRead YESNO NOT NULL,
    RelatedAppointmentID LONG
);
'@,
@'
CREATE TABLE VaccineAllocation (
    AllocationID AUTOINCREMENT PRIMARY KEY,
    BatchID LONG NOT NULL,
    CenterID LONG NOT NULL,
    QuantityAllocated LONG NOT NULL,
    AllocationDate DATETIME NOT NULL,
    AllocatedBy LONG NOT NULL
);
'@,
@'
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
'@,
@'
CREATE TABLE AuditLog (
    LogID AUTOINCREMENT PRIMARY KEY,
    UserID LONG NOT NULL,
    [Action] TEXT(20) NOT NULL,
    TableName TEXT(50) NOT NULL,
    RecordID LONG NOT NULL,
    ActionDate DATETIME NOT NULL,
    Details TEXT(255)
);
'@
    )
}

function Build-IndexSql() {
    return @(
        "CREATE UNIQUE INDEX idx_Citizens_HKID ON Citizens (HKID);",
        "CREATE UNIQUE INDEX idx_Users_Username ON Users (Username);",
        "CREATE UNIQUE INDEX idx_Batches_BatchNumber ON VaccineBatches (BatchNumber);",
        "CREATE INDEX idx_Appointments_CitizenID ON Appointments (CitizenID);",
        "CREATE INDEX idx_Appointments_SlotID ON Appointments (SlotID);",
        "CREATE INDEX idx_VaccinationRecords_AppointmentID ON VaccinationRecords (AppointmentID);",
        "CREATE INDEX idx_VaccinationRecords_CitizenID ON VaccinationRecords (CitizenID);",
        "CREATE INDEX idx_Screening_AppointmentID ON PreVaccinationScreening (AppointmentID);",
        "CREATE INDEX idx_Consent_AppointmentID ON ConsentRecords (AppointmentID);",
        "CREATE INDEX idx_Observation_RecordID ON ObservationRecords (RecordID);",
        "CREATE INDEX idx_SideEffects_RecordID ON SideEffects (VaccinationRecordID);",
        "CREATE INDEX idx_FollowUps_ReportID ON AdverseEffectFollowUps (ReportID);",
        "CREATE INDEX idx_Notifications_CitizenID ON Notifications (CitizenID);",
        "CREATE INDEX idx_Inventory_BatchID ON InventoryTransactions (BatchID);"
    )
}

function Get-VaccineRows() {
    return @(
        @{ VaccineName='Comirnaty Adult'; Manufacturer='Pfizer-BioNTech'; RequiredDoses=2; DoseIntervalDays=21; StorageTemp='2-8C'; MinAge=12; MaxAge=120 },
        @{ VaccineName='Comirnaty Bivalent'; Manufacturer='Pfizer-BioNTech'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=12; MaxAge=120 },
        @{ VaccineName='Moderna Adult'; Manufacturer='Moderna'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='Moderna Booster'; Manufacturer='Moderna'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='Sinovac Standard'; Manufacturer='Sinovac'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='Novavax'; Manufacturer='Novavax'; RequiredDoses=2; DoseIntervalDays=21; StorageTemp='2-8C'; MinAge=12; MaxAge=120 },
        @{ VaccineName='Vaxigrip Tetra'; Manufacturer='Sanofi'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=6; MaxAge=120 },
        @{ VaccineName='Influvac Tetra'; Manufacturer='Abbott'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=6; MaxAge=120 },
        @{ VaccineName='Fluarix Tetra'; Manufacturer='GSK'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=3; MaxAge=120 },
        @{ VaccineName='Gardasil 9'; Manufacturer='MSD'; RequiredDoses=3; DoseIntervalDays=60; StorageTemp='2-8C'; MinAge=9; MaxAge=45 },
        @{ VaccineName='Cervarix'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=180; StorageTemp='2-8C'; MinAge=9; MaxAge=45 },
        @{ VaccineName='Prevenar 13'; Manufacturer='Pfizer'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=65; MaxAge=120 },
        @{ VaccineName='Pneumovax 23'; Manufacturer='MSD'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=65; MaxAge=120 },
        @{ VaccineName='Shingrix'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=60; StorageTemp='2-8C'; MinAge=50; MaxAge=120 },
        @{ VaccineName='Engerix-B Adult'; Manufacturer='GSK'; RequiredDoses=3; DoseIntervalDays=30; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='Havrix Adult'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=180; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='Twinrix Adult'; Manufacturer='GSK'; RequiredDoses=3; DoseIntervalDays=30; StorageTemp='2-8C'; MinAge=18; MaxAge=120 },
        @{ VaccineName='MMR II'; Manufacturer='MSD'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=1; MaxAge=50 },
        @{ VaccineName='Priorix'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=1; MaxAge=50 },
        @{ VaccineName='Varivax'; Manufacturer='MSD'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=1; MaxAge=50 },
        @{ VaccineName='Bexsero'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=30; StorageTemp='2-8C'; MinAge=2; MaxAge=50 },
        @{ VaccineName='Menveo'; Manufacturer='GSK'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=2; MaxAge=55 },
        @{ VaccineName='Boostrix'; Manufacturer='GSK'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=10; MaxAge=120 },
        @{ VaccineName='Adacel'; Manufacturer='Sanofi'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=10; MaxAge=120 },
        @{ VaccineName='Rotarix'; Manufacturer='GSK'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=0; MaxAge=1 },
        @{ VaccineName='Infanrix Hexa'; Manufacturer='GSK'; RequiredDoses=3; DoseIntervalDays=30; StorageTemp='2-8C'; MinAge=0; MaxAge=2 },
        @{ VaccineName='ActHIB'; Manufacturer='Sanofi'; RequiredDoses=3; DoseIntervalDays=30; StorageTemp='2-8C'; MinAge=0; MaxAge=5 },
        @{ VaccineName='Arexvy'; Manufacturer='GSK'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=60; MaxAge=120 },
        @{ VaccineName='Abrysvo'; Manufacturer='Pfizer'; RequiredDoses=1; DoseIntervalDays=0; StorageTemp='2-8C'; MinAge=60; MaxAge=120 },
        @{ VaccineName='Ixiaro'; Manufacturer='Valneva'; RequiredDoses=2; DoseIntervalDays=28; StorageTemp='2-8C'; MinAge=2; MaxAge=65 }
    )
}

function Get-CenterRows() {
    $districts = @(
        'Central','Admiralty','Wan Chai','Causeway Bay','North Point','Quarry Bay','Aberdeen','Pok Fu Lam','Tsim Sha Tsui','Jordan',
        'Yau Ma Tei','Mong Kok','Sham Shui Po','Cheung Sha Wan','Kwun Tong','Lam Tin','Sai Kung','Sha Tin','Tai Po','Tuen Mun',
        'Tsuen Wan','Kwai Chung','Yuen Long','Tin Shui Wai','Tseung Kwan O','Fanling','Sheung Shui','Ma On Shan','Discovery Bay','Lantau'
    )
    $rows = @()
    for ($i = 0; $i -lt $districts.Count; $i++) {
        $rows += @{
            CenterName = $districts[$i] + ' Community Vaccination Centre'
            Address = ($i + 1).ToString() + ' Health Road'
            Region = $(if ($i -lt 8) { 'Hong Kong Island' } elseif ($i -lt 17) { 'Kowloon' } else { 'New Territories' })
            ContactPhone = '28' + (100000 + $i).ToString().PadLeft(6,'0')
            OperatingHours = $(if ($i % 2 -eq 0) { '08:30-17:30' } else { '09:00-18:00' })
            MaxDailyService = 80 + (($i % 6) * 10)
            DoctorName = 'Dr. ' + @('Adams','Baker','Clark','Davis','Evans','Foster','Green','Hall','Irwin','Jones')[$i % 10]
            IsActive = $true
        }
    }
    return $rows
}

function Get-CitizenRows() {
    $firstNames = @('Ethan','Noah','Liam','Mason','Lucas','Aiden','Owen','Leo','Ryan','Caleb','Nathan','Dylan','Isaac','Aaron','Evan','Wyatt','Julian','Oscar','Adam','Jason','Chloe','Emma','Ava','Mia','Ella','Grace','Lily','Zoe','Nora','Ruby','Sophie','Ivy','Clara','Anna','Lucy','Hannah','Stella','Naomi','Hazel','Alice')
    $lastNames = @('Carter','Bennett','Foster','Parker','Hayes','Brooks','Cooper','Morgan','Bell','Ward','Reed','Cook','Bailey','Price','Long','Powell','Ross','Gray','James','Kelly','Scott','Rivera','Murphy','Sanders','Russell','Cole','Jenkins','Perry','Butler','Sullivan','Bryant','Griffin','Hayden','West','Mills','Ford','Wells','Barnes','Stone','Fisher')
    $districts = @('Hong Kong Island','Kowloon','New Territories')
    $allergies = @('None','Penicillin','Egg','Latex','Seafood','Aspirin','Shellfish','Dust')
    $occupations = @('Student','Teacher','Engineer','Nurse','Designer','Analyst','Chef','Driver','Office Clerk','Retired','Sales Assistant','Developer')
    $birthPlaces = @('Hong Kong','Macau','Shenzhen','Guangzhou','Singapore','London')
    $agePlan = @(7,9,10,6,8,11,13,14,15,16,17,12,19,21,23,24,26,27,28,29,31,33,35,36,38,40,42,44,46,48,52,55,58,59,61,64,67,70,74,78)
    $rows = @()
    for ($i = 0; $i -lt 40; $i++) {
        $first = $firstNames[$i]
        $last = $lastNames[$i]
        $dob = Get-DobFromAge $agePlan[$i] ($i % 20)
        $rows += @{
            HKID = [string]([char](65 + ($i % 26))) + (100000 + $i).ToString() + '(' + (($i % 9) + 1).ToString() + ')'
            CitizenName = $first + ' ' + $last
            FirstName = $first
            LastName = $last
            Gender = $(if ($i -lt 20) { 'Male' } else { 'Female' })
            DOB = $dob
            Phone = '91' + (200000 + $i).ToString().PadLeft(6,'0')
            Email = ($first + '.' + $last + '@example.com').ToLower()
            Address = ($i + 10).ToString() + ' Sample Street'
            District = $districts[$i % 3]
            AllergyHistory = $allergies[$i % $allergies.Count]
            Occupation = $occupations[$i % $occupations.Count]
            PlaceOfBirth = $birthPlaces[$i % $birthPlaces.Count]
        }
    }
    return $rows
}

function Get-UserRows() {
    $rows = @(
        @{ Username='admin01'; UserPassword='adminpass'; Role='Admin'; LinkedCitizenID=$null; IsActive=$true },
        @{ Username='admin02'; UserPassword='adminpass'; Role='Admin'; LinkedCitizenID=$null; IsActive=$true },
        @{ Username='it_support01'; UserPassword='pass123'; Role='ITSupport'; LinkedCitizenID=$null; IsActive=$true },
        @{ Username='it_support02'; UserPassword='pass123'; Role='ITSupport'; LinkedCitizenID=$null; IsActive=$true }
    )
    for ($i = 1; $i -le 12; $i++) {
        $rows += @{ Username=('nurse_' + $i.ToString('00')); UserPassword='pass123'; Role='MedicalStaff'; LinkedCitizenID=$null; IsActive=$true }
    }
    for ($i = 1; $i -le 14; $i++) {
        $rows += @{ Username=('citizen_' + $i.ToString('00')); UserPassword='pass123'; Role='Citizen'; LinkedCitizenID=$i; IsActive=$true }
    }
    return $rows
}

function Get-BatchRows($vaccineRows) {
    $suppliers = @('HK Health Supply','MediChain Asia','VaxLink Logistics','CarePlus Medical','North Harbor Pharma')
    $rows = @()
    for ($i = 1; $i -le 40; $i++) {
        $vaccineId = (($i - 1) % 30) + 1
        $manufactureDate = $today.AddMonths(-8 + ($i % 4)).AddDays(-($i * 2))
        if ($i -le 5) {
            $expiryDate = $today.AddDays(8 + ($i * 3))
            $currentStock = 6 + $i
        } elseif ($i -le 10) {
            $expiryDate = $today.AddDays(20 + $i)
            $currentStock = 18 + $i
        } else {
            $expiryDate = $today.AddMonths(4 + ($i % 6)).AddDays($i)
            $currentStock = 60 + (($i % 9) * 12)
        }
        $rows += @{
            BatchNumber = 'BAT-2026-' + $i.ToString('000')
            VaccineID = $vaccineId
            ManufactureDate = $manufactureDate
            ExpiryDate = $expiryDate
            CurrentStock = $currentStock
            Supplier = $suppliers[$i % $suppliers.Count]
        }
    }
    return $rows
}

function New-SlotRow([int]$centerId, [int]$vaccineId, [datetime]$slotDate, [int]$capacity) {
    return @{
        CenterID = $centerId
        VaccineID = $vaccineId
        SlotDate = $slotDate.Date
        SlotTime = $slotDate
        Capacity = $capacity
        BookedCount = 0
        SlotStatus = 'Open'
    }
}

function Get-AppointmentSlotRows() {
    $rows = @()

    for ($i = 0; $i -lt 10; $i++) {
        $rows += New-SlotRow ($i + 1) 1 ((Get-Date '2026-01-05 09:00').AddDays($i)) $(if ($i % 2 -eq 0) { 2 } else { 4 })
    }
    for ($i = 0; $i -lt 5; $i++) {
        $rows += New-SlotRow ($i + 11) 3 ((Get-Date '2026-01-16 10:00').AddDays($i)) 3
    }
    for ($i = 0; $i -lt 5; $i++) {
        $rows += New-SlotRow ($i + 11) 3 ((Get-Date '2026-02-15 10:00').AddDays($i)) $(if ($i % 2 -eq 0) { 2 } else { 3 })
    }
    for ($i = 0; $i -lt 5; $i++) {
        $rows += New-SlotRow ($i + 16) 5 ((Get-Date '2026-03-10 11:00').AddDays($i)) 4
    }
    for ($i = 0; $i -lt 5; $i++) {
        $rows += New-SlotRow ($i + 21) 7 ((Get-Date '2026-03-01 14:00').AddDays($i)) 5
    }
    $rows += New-SlotRow 26 6 (Get-Date '2026-01-08 09:30') 2
    $rows += New-SlotRow 27 6 (Get-Date '2026-01-09 09:30') 2
    $rows += New-SlotRow 26 6 (Get-Date '2026-01-31 09:30') 2
    $rows += New-SlotRow 27 6 (Get-Date '2026-02-01 09:30') 2
    $rows += New-SlotRow 28 14 (Get-Date '2026-01-20 15:00') 2
    $rows += New-SlotRow 29 14 (Get-Date '2026-01-21 15:00') 2
    $rows += New-SlotRow 4 1 (Get-Date '2026-04-02 09:00') 1
    $rows += New-SlotRow 9 1 (Get-Date '2026-04-02 09:00') 1
    $rows += New-SlotRow 6 5 (Get-Date '2026-04-04 14:00') 2
    $rows += New-SlotRow 12 5 (Get-Date '2026-04-04 14:00') 2

    $extraVaccines = @(1,3,5,7,8,10,12,14,18,21,22,23,24,28,29,30,6,9,13,15)
    for ($i = 0; $i -lt 20; $i++) {
        $centerId = (($i + 9) % 30) + 1
        $dateTime = (Get-Date '2026-04-10 09:00').AddDays([math]::Floor($i / 2)).AddHours($(if ($i % 2 -eq 0) { 0 } else { 5 }))
        $rows += New-SlotRow $centerId $extraVaccines[$i] $dateTime (3 + ($i % 4))
    }

    return $rows
}

function Get-AppointmentRows($slotRows) {
    $rows = @()

    for ($i = 1; $i -le 10; $i++) {
        $slot = $slotRows[$i - 1]
        $rows += @{
            CitizenID = $i
            VaccineID = $slot.VaccineID
            CenterID = $slot.CenterID
            SlotID = $i
            BatchID = (Resolve-BatchId $slot.VaccineID 1)
            AppointmentDate = $slot.SlotTime
            DoseNo = 1
            AppointmentStatus = 'Completed'
            ActualVaccinationDate = $slot.SlotTime.AddMinutes(20)
            VaccinatedBy = 5 + (($i - 1) % 12)
            Notes = 'Completed first dose; due for follow-up reminder.'
            CancellationReason = $null
        }
    }

    for ($i = 11; $i -le 15; $i++) {
        $slot = $slotRows[$i - 1]
        $rows += @{
            CitizenID = $i
            VaccineID = $slot.VaccineID
            CenterID = $slot.CenterID
            SlotID = $i
            BatchID = (Resolve-BatchId $slot.VaccineID 1)
            AppointmentDate = $slot.SlotTime
            DoseNo = 1
            AppointmentStatus = 'Completed'
            ActualVaccinationDate = $slot.SlotTime.AddMinutes(15)
            VaccinatedBy = 5 + (($i - 1) % 12)
            Notes = 'Completed first dose; returned later for second dose.'
            CancellationReason = $null
        }
    }

    for ($i = 16; $i -le 20; $i++) {
        $slot = $slotRows[$i - 1]
        $citizenId = $i - 5
        $rows += @{
            CitizenID = $citizenId
            VaccineID = $slot.VaccineID
            CenterID = $slot.CenterID
            SlotID = $i
            BatchID = (Resolve-BatchId $slot.VaccineID 2)
            AppointmentDate = $slot.SlotTime
            DoseNo = 2
            AppointmentStatus = 'Completed'
            ActualVaccinationDate = $slot.SlotTime.AddMinutes(18)
            VaccinatedBy = 5 + (($i - 1) % 12)
            Notes = 'Series completed successfully.'
            CancellationReason = $null
        }
    }

    for ($i = 21; $i -le 25; $i++) {
        $slot = $slotRows[$i - 1]
        $citizenId = $i - 5
        $rows += @{
            CitizenID = $citizenId
            VaccineID = $slot.VaccineID
            CenterID = $slot.CenterID
            SlotID = $i
            BatchID = (Resolve-BatchId $slot.VaccineID 1)
            AppointmentDate = $slot.SlotTime
            DoseNo = 1
            AppointmentStatus = 'Completed'
            ActualVaccinationDate = $slot.SlotTime.AddMinutes(12)
            VaccinatedBy = 5 + (($i - 1) % 12)
            Notes = 'Recent first dose; second dose not due yet.'
            CancellationReason = $null
        }
    }

    for ($i = 26; $i -le 30; $i++) {
        $slot = $slotRows[$i - 1]
        $citizenId = $i - 5
        $rows += @{
            CitizenID = $citizenId
            VaccineID = $slot.VaccineID
            CenterID = $slot.CenterID
            SlotID = $i
            BatchID = (Resolve-BatchId $slot.VaccineID 1)
            AppointmentDate = $slot.SlotTime
            DoseNo = 1
            AppointmentStatus = 'Completed'
            ActualVaccinationDate = $slot.SlotTime.AddMinutes(10)
            VaccinatedBy = 5 + (($i - 1) % 12)
            Notes = 'Seasonal vaccine administered.'
            CancellationReason = $null
        }
    }

    $rows += @{
        CitizenID = 26
        VaccineID = 6
        CenterID = 26
        SlotID = 31
        BatchID = (Resolve-BatchId 6 1)
        AppointmentDate = $slotRows[30].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[30].SlotTime.AddMinutes(20)
        VaccinatedBy = 8
        Notes = 'Novavax first dose.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 27
        VaccineID = 6
        CenterID = 27
        SlotID = 32
        BatchID = (Resolve-BatchId 6 1)
        AppointmentDate = $slotRows[31].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[31].SlotTime.AddMinutes(20)
        VaccinatedBy = 9
        Notes = 'Novavax first dose.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 26
        VaccineID = 6
        CenterID = 26
        SlotID = 33
        BatchID = (Resolve-BatchId 6 2)
        AppointmentDate = $slotRows[32].SlotTime
        DoseNo = 2
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[32].SlotTime.AddMinutes(15)
        VaccinatedBy = 10
        Notes = 'Novavax series completed.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 27
        VaccineID = 6
        CenterID = 27
        SlotID = 34
        BatchID = (Resolve-BatchId 6 2)
        AppointmentDate = $slotRows[33].SlotTime
        DoseNo = 2
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[33].SlotTime.AddMinutes(15)
        VaccinatedBy = 11
        Notes = 'Novavax series completed.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 35
        VaccineID = 14
        CenterID = 28
        SlotID = 35
        BatchID = (Resolve-BatchId 14 1)
        AppointmentDate = $slotRows[34].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[34].SlotTime.AddMinutes(15)
        VaccinatedBy = 12
        Notes = 'Shingrix first dose; follow-up due later.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 29
        VaccineID = 14
        CenterID = 29
        SlotID = 36
        BatchID = (Resolve-BatchId 14 1)
        AppointmentDate = $slotRows[35].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Completed'
        ActualVaccinationDate = $slotRows[35].SlotTime.AddMinutes(15)
        VaccinatedBy = 13
        Notes = 'Shingrix first dose; follow-up due later.'
        CancellationReason = $null
    }

    $rows += @{
        CitizenID = 30
        VaccineID = 1
        CenterID = 4
        SlotID = 37
        BatchID = $null
        AppointmentDate = $slotRows[36].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Scheduled'
        ActualVaccinationDate = $null
        VaccinatedBy = $null
        Notes = 'Duplicate booking to test conflict group A.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 30
        VaccineID = 1
        CenterID = 9
        SlotID = 38
        BatchID = $null
        AppointmentDate = $slotRows[37].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Scheduled'
        ActualVaccinationDate = $null
        VaccinatedBy = $null
        Notes = 'Duplicate booking to test conflict group A.'
        CancellationReason = $null
    }
    $rows += @{
        CitizenID = 31
        VaccineID = 5
        CenterID = 6
        SlotID = 39
        BatchID = $null
        AppointmentDate = $slotRows[38].SlotTime
        DoseNo = 1
        AppointmentStatus = 'Cancelled'
        ActualVaccinationDate = $null
        VaccinatedBy = $null
        Notes = 'Cancelled duplicate conflict case.'
        CancellationReason = 'Fever before arrival'
    }
    $rows += @{
        CitizenID = 31
        VaccineID = 5
        CenterID = 12
        SlotID = 40
        BatchID = $null
        AppointmentDate = $slotRows[39].SlotTime
        DoseNo = 1
        AppointmentStatus = 'NoShow'
        ActualVaccinationDate = $null
        VaccinatedBy = $null
        Notes = 'No-show duplicate conflict case.'
        CancellationReason = $null
    }

    return $rows
}

function Update-SlotUsage($slotRows, $appointmentRows) {
    $counts = @{}
    foreach ($appointment in $appointmentRows) {
        if ($appointment.AppointmentStatus -ne 'Cancelled') {
            $slotId = [int]$appointment.SlotID
            if (-not $counts.ContainsKey($slotId)) { $counts[$slotId] = 0 }
            $counts[$slotId]++
        }
    }

    for ($i = 0; $i -lt $slotRows.Count; $i++) {
        $slotId = $i + 1
        $booked = $(if ($counts.ContainsKey($slotId)) { $counts[$slotId] } else { 0 })
        $slotRows[$i].BookedCount = $booked
        if ($slotRows[$i].SlotTime -lt $today -and $booked -gt 0) {
            $slotRows[$i].SlotStatus = 'Served'
        } elseif ($booked -ge $slotRows[$i].Capacity) {
            $slotRows[$i].SlotStatus = 'Full'
        } else {
            $slotRows[$i].SlotStatus = 'Open'
        }
    }
}

function Get-VaccinationRecordRows($appointmentRows) {
    $rows = @()
    for ($i = 0; $i -lt $appointmentRows.Count; $i++) {
        $appointment = $appointmentRows[$i]
        if ($appointment.AppointmentStatus -eq 'Completed') {
            $rows += @{
                AppointmentID = $i + 1
                CitizenID = $appointment.CitizenID
                VaccineID = $appointment.VaccineID
                BatchID = $appointment.BatchID
                DoseNo = $appointment.DoseNo
                VaccinationDate = $appointment.ActualVaccinationDate
                StaffUserID = $appointment.VaccinatedBy
                OutcomeStatus = $(if (($i + 1) % 9 -eq 0) { 'Completed with advice' } else { 'Completed' })
            }
        }
    }
    return $rows
}

function Get-ScreeningRows($appointmentRows, $citizenRows) {
    $rows = @()
    for ($i = 0; $i -lt $appointmentRows.Count; $i++) {
        $appointment = $appointmentRows[$i]
        $citizen = $citizenRows[$appointment.CitizenID - 1]
        $riskFlag = (($i + 1) % 7 -eq 0) -or ($appointment.AppointmentStatus -eq 'Cancelled')
        $temp = $(if ($riskFlag) { 37.5 + (($i % 3) * 0.2) } else { 36.4 + (($i % 5) * 0.2) })
        $allergyConfirmed = ($citizen.AllergyHistory -ne 'None')
        $note = $null
        if ($appointment.AppointmentStatus -eq 'Cancelled') {
            $note = 'Citizen reported fever and appointment was cancelled.'
        } elseif ($riskFlag) {
            $note = 'Flagged for nurse review before proceeding.'
        } elseif ($allergyConfirmed) {
            $note = 'Known allergy reviewed before vaccination.'
        }
        $rows += @{
            AppointmentID = $i + 1
            CitizenID = $appointment.CitizenID
            ScreeningDate = $appointment.AppointmentDate.AddMinutes(-30)
            RiskFlag = $riskFlag
            Temperature = [math]::Round($temp, 1)
            AllergyConfirmed = $allergyConfirmed
            ContraindicationNote = $note
            ScreenedBy = 5 + ($i % 12)
        }
    }
    return $rows
}

function Get-ConsentRows($appointmentRows) {
    $rows = @()
    for ($i = 0; $i -lt $appointmentRows.Count; $i++) {
        $appointment = $appointmentRows[$i]
        $status = 'Granted'
        $note = 'Citizen consent captured before clinical workflow.'
        if ($appointment.AppointmentStatus -eq 'Cancelled') {
            $status = 'Declined'
            $note = 'Citizen declined on day due to fever symptoms.'
        } elseif ($appointment.AppointmentStatus -eq 'Scheduled') {
            $status = 'Pending'
            $note = 'Consent will be captured on arrival.'
        }
        $rows += @{
            AppointmentID = $i + 1
            CitizenID = $appointment.CitizenID
            ConsentStatus = $status
            ConsentDate = $appointment.AppointmentDate.AddMinutes(-10)
            WitnessedBy = 5 + ($i % 12)
            ConsentNote = $note
        }
    }
    return $rows
}

function Get-ObservationRows($vaccinationRecordRows) {
    $rows = @()
    for ($i = 0; $i -lt $vaccinationRecordRows.Count; $i++) {
        $record = $vaccinationRecordRows[$i]
        $immediate = ((($i + 1) % 6) -eq 0) -or (($i + 1) -in @(3, 14))
        $outcome = $(if ($immediate) { $(if (($i + 1) % 5 -eq 0) { 'Referred to clinic' } else { 'Extended observation' }) } else { 'Stable discharge' })
        $rows += @{
            RecordID = $i + 1
            CitizenID = $record.CitizenID
            ObservationStart = $record.VaccinationDate
            ObservationEnd = $record.VaccinationDate.AddMinutes($(if ($immediate) { 30 } else { 15 }))
            ImmediateReaction = $immediate
            ObservationOutcome = $outcome
            ObservedBy = 5 + (($i + 2) % 12)
        }
    }
    return $rows
}

function Get-SideEffectRows($vaccinationRecordRows) {
    $mildSymptoms = @('Injection-site pain','Fatigue','Headache','Low-grade fever','Dizziness')
    $moderateSymptoms = @('Fever above 38C','Arm swelling','Nausea','Body ache','Vomiting')
    $severeSymptoms = @('Chest tightness','Shortness of breath','Persistent high fever','Rapid heartbeat','Allergic rash')
    $rows = @()
    for ($i = 0; $i -lt 35; $i++) {
        $record = $vaccinationRecordRows[$i]
        $severity = $(if (($i + 1) % 11 -eq 0) { 'Severe' } elseif (($i + 1) % 3 -eq 0) { 'Moderate' } else { 'Mild' })
        $symptom = $(if ($severity -eq 'Severe') { $severeSymptoms[$i % $severeSymptoms.Count] } elseif ($severity -eq 'Moderate') { $moderateSymptoms[$i % $moderateSymptoms.Count] } else { $mildSymptoms[$i % $mildSymptoms.Count] })
        $status = $(if ($severity -eq 'Severe') { 'Escalated' } elseif ($severity -eq 'Moderate') { 'Monitoring' } else { 'Closed' })
        $rows += @{
            VaccinationRecordID = $i + 1
            AppointmentID = $record.AppointmentID
            CitizenID = $record.CitizenID
            SymptomType = $symptom
            SeverityLevel = $severity
            ReportDate = $record.VaccinationDate.AddDays(($i % 4) + 1)
            Details = 'Reported via follow-up survey within one week of vaccination.'
            FollowUpStatus = $status
        }
    }
    return $rows
}

function Get-FollowUpRows($sideEffectRows) {
    $rows = @()
    for ($i = 0; $i -lt 30; $i++) {
        $report = $sideEffectRows[$i]
        $outcome = $(if ($report.SeverityLevel -eq 'Severe') { $(if (($i + 1) % 2 -eq 0) { 'Needs Review' } else { 'Open' }) } elseif ($report.SeverityLevel -eq 'Moderate') { $(if (($i + 1) % 4 -eq 0) { 'Monitoring' } else { 'Recovered' }) } else { 'Recovered' })
        $rows += @{
            ReportID = $i + 1
            FollowUpDate = $report.ReportDate.AddDays(1)
            StaffUserID = 5 + ($i % 12)
            ActionTaken = $(if ($report.SeverityLevel -eq 'Severe') { 'Escalated to clinician and arranged same-day review.' } elseif ($report.SeverityLevel -eq 'Moderate') { 'Phone review completed and home-care advice provided.' } else { 'Automated check-in closed after recovery confirmed.' })
            FollowUpOutcome = $outcome
            EscalationFlag = ($report.SeverityLevel -eq 'Severe')
        }
    }
    return $rows
}

function Get-NotificationRows($appointmentRows) {
    $rows = @()
    for ($i = 0; $i -lt 15; $i++) {
        $rows += @{
            CitizenID = $i + 1
            NotificationType = 'SecondDose'
            Message = 'Second dose is now due based on your first vaccination date.'
            SendDate = $today.AddDays(-($i % 5))
            IsRead = ($i % 3 -eq 0)
            RelatedAppointmentID = $i + 1
        }
    }
    for ($i = 15; $i -lt 25; $i++) {
        $rows += @{
            CitizenID = (($i - 15) % 10) + 1
            NotificationType = 'SideEffectFollowUp'
            Message = 'A nurse follow-up has been scheduled for your reported symptoms.'
            SendDate = $today.AddDays(-($i % 4))
            IsRead = ($i % 2 -eq 0)
            RelatedAppointmentID = $i - 14
        }
    }
    for ($i = 25; $i -lt 40; $i++) {
        $rows += @{
            CitizenID = (($i - 25) % 15) + 16
            NotificationType = 'ExpiryAlert'
            Message = 'Operational alert record created for stock monitoring demonstration.'
            SendDate = $today.AddDays(-($i % 6))
            IsRead = $false
            RelatedAppointmentID = $null
        }
    }
    return $rows
}

function Get-AllocationRows() {
    $rows = @()
    for ($i = 1; $i -le 40; $i++) {
        $rows += @{
            BatchID = $i
            CenterID = (($i - 1) % 30) + 1
            QuantityAllocated = 40 + (($i % 8) * 10)
            AllocationDate = $today.AddDays(-70 + $i)
            AllocatedBy = $(if ($i % 2 -eq 0) { 1 } else { 2 })
        }
    }
    return $rows
}

function Get-InventoryTransactionRows($vaccinationRecordRows) {
    $rows = @()
    for ($i = 1; $i -le 15; $i++) {
        $rows += @{
            BatchID = $i
            CenterID = (($i - 1) % 30) + 1
            TransactionType = 'Allocation'
            Quantity = 30 + (($i % 5) * 5)
            TransactionDate = $today.AddDays(-60 + $i)
            RelatedRecordID = $null
            ProcessedBy = 1
        }
    }
    for ($i = 1; $i -le 15; $i++) {
        $rows += @{
            BatchID = Resolve-BatchId $vaccinationRecordRows[$i - 1].VaccineID $vaccinationRecordRows[$i - 1].DoseNo
            CenterID = ((($i - 1) % 15) + 1)
            TransactionType = 'Used'
            Quantity = 1
            TransactionDate = $vaccinationRecordRows[$i - 1].VaccinationDate
            RelatedRecordID = $i
            ProcessedBy = 5 + (($i - 1) % 12)
        }
    }
    for ($i = 1; $i -le 5; $i++) {
        $rows += @{
            BatchID = $i
            CenterID = $i
            TransactionType = 'Wastage'
            Quantity = 1 + ($i % 3)
            TransactionDate = $today.AddDays(-10 + $i)
            RelatedRecordID = $null
            ProcessedBy = 2
        }
    }
    for ($i = 1; $i -le 5; $i++) {
        $rows += @{
            BatchID = $i + 5
            CenterID = $i + 5
            TransactionType = 'Adjustment'
            Quantity = 1 + ($i % 2)
            TransactionDate = $today.AddDays(-5 + $i)
            RelatedRecordID = $null
            ProcessedBy = 2
        }
    }
    return $rows
}

function Get-AuditRows() {
    $tables = @('Citizens','Appointments','VaccinationRecords','PreVaccinationScreening','ConsentRecords','ObservationRecords','SideEffects','InventoryTransactions')
    $actions = @('INSERT','UPDATE','INSERT','UPDATE','INSERT')
    $rows = @()
    for ($i = 1; $i -le 40; $i++) {
        $rows += @{
            UserID = $(if ($i % 3 -eq 0) { 1 } else { 5 + (($i - 1) % 12) })
            Action = $actions[$i % $actions.Count]
            TableName = $tables[$i % $tables.Count]
            RecordID = $i
            ActionDate = $today.AddDays(-40 + $i).AddHours($i % 8)
            Details = 'Sample audit trail entry for v2 workflow verification.'
        }
    }
    return $rows
}

function Build-QuerySqlMap() {
    $ageExpr = 'DateDiff("yyyy",[DOB],Date()) - IIf(Format([DOB],"mmdd")>Format(Date(),"mmdd"),1,0)'
    $ageGroupExpr = 'IIf(' + $ageExpr + '<=11,"0-11",IIf(' + $ageExpr + '<=17,"12-17",IIf(' + $ageExpr + '<=29,"18-29",IIf(' + $ageExpr + '<=44,"30-44",IIf(' + $ageExpr + '<=59,"45-59","60+")))))'
    $map = [ordered]@{}

    $map['qry_UI_CitizenProfile'] = @"
SELECT
    Citizens.CitizenID,
    Citizens.HKID,
    Citizens.CitizenName,
    Citizens.FirstName,
    Citizens.LastName,
    Citizens.Gender,
    Citizens.DOB,
    $ageExpr AS CitizenAge,
    $ageGroupExpr AS AgeGroup,
    Citizens.Phone,
    Citizens.Email,
    Citizens.[Address],
    Citizens.District,
    Citizens.Occupation,
    Citizens.PlaceOfBirth,
    Citizens.AllergyHistory
FROM Citizens
ORDER BY Citizens.CitizenID;
"@

    $map['qry_UI_Appointments'] = @"
SELECT
    A.AppointmentID,
    A.CitizenID,
    C.HKID,
    C.CitizenName,
    A.VaccineID,
    V.VaccineName,
    A.CenterID,
    VC.CenterName,
    A.SlotID,
    S.SlotDate,
    Format([S].[SlotTime],'hh:nn AM/PM') AS SlotTimeLabel,
    A.BatchID,
    B.BatchNumber,
    A.DoseNo,
    A.AppointmentDate,
    A.AppointmentStatus,
    A.ActualVaccinationDate,
    A.VaccinatedBy,
    U.Username AS VaccinatedByName,
    A.Notes,
    A.CancellationReason
FROM ((((((Appointments AS A
INNER JOIN Citizens AS C ON A.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON A.VaccineID = V.VaccineID)
INNER JOIN VaccinationCenters AS VC ON A.CenterID = VC.CenterID)
LEFT JOIN AppointmentSlots AS S ON A.SlotID = S.SlotID)
LEFT JOIN VaccineBatches AS B ON A.BatchID = B.BatchID)
LEFT JOIN Users AS U ON A.VaccinatedBy = U.UserID)
ORDER BY A.AppointmentDate DESC, A.AppointmentID DESC;
"@

    $map['qry_AppointmentConflictCheck'] = @"
SELECT
    A.CitizenID,
    C.CitizenName,
    DateValue([A].[AppointmentDate]) AS ConflictDate,
    Count(A.AppointmentID) AS ConflictCount,
    Min(A.AppointmentID) AS FirstAppointmentID,
    Max(A.AppointmentID) AS LastAppointmentID
FROM Appointments AS A
INNER JOIN Citizens AS C ON A.CitizenID = C.CitizenID
WHERE A.AppointmentStatus <> 'Completed'
GROUP BY A.CitizenID, C.CitizenName, DateValue([A].[AppointmentDate])
HAVING Count(A.AppointmentID) > 1
ORDER BY Count(A.AppointmentID) DESC, DateValue([A].[AppointmentDate]) DESC;
"@

    $map['qry_SecondDoseReminder'] = @"
SELECT
    VR.CitizenID,
    C.CitizenName,
    V.VaccineName,
    VR.VaccinationDate AS FirstDoseDate,
    DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate) AS RecommendedSecondDoseDate,
    DateDiff('d', DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate), Date()) AS DaysOverdue
FROM (VaccinationRecords AS VR
INNER JOIN Citizens AS C ON VR.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID
WHERE VR.DoseNo = 1
  AND V.RequiredDoses >= 2
  AND DateAdd('d', V.DoseIntervalDays, VR.VaccinationDate) <= Date()
  AND DCount('*','VaccinationRecords','CitizenID=' & [VR].[CitizenID] & ' AND VaccineID=' & [VR].[VaccineID] & ' AND DoseNo>=2') = 0
ORDER BY 6 DESC, 2;
"@

    $map['qry_ExpiringOrLowStockBatches'] = @"
SELECT
    B.BatchID,
    B.BatchNumber,
    V.VaccineName,
    B.ExpiryDate,
    B.CurrentStock,
    B.Supplier,
    IIf([CurrentStock] < 20, 'Low Stock', 'OK') AS StockFlag,
    IIf(DateDiff('d', Date(), [ExpiryDate]) <= 30, 'Expiring Soon', 'OK') AS ExpiryFlag
FROM VaccineBatches AS B
INNER JOIN Vaccines AS V ON B.VaccineID = V.VaccineID
WHERE B.CurrentStock < 20
   OR DateDiff('d', Date(), [ExpiryDate]) <= 30
ORDER BY B.ExpiryDate, B.CurrentStock;
"@

    $map['qry_CenterCapacityUsage'] = @"
SELECT
    VC.CenterName,
    S.SlotDate,
    Format([S].[SlotTime],'hh:nn AM/PM') AS SlotTimeLabel,
    S.Capacity,
    S.BookedCount,
    IIf([S].[Capacity]=0,0,([S].[BookedCount]*100.0)/[S].[Capacity]) AS UtilizationPercent,
    S.SlotStatus
FROM AppointmentSlots AS S
INNER JOIN VaccinationCenters AS VC ON S.CenterID = VC.CenterID
ORDER BY S.SlotDate, VC.CenterName, S.SlotTime;
"@

    $map['qry_CitizenVaccinationHistory'] = @"
SELECT
    C.CitizenID,
    C.CitizenName,
    A.AppointmentID,
    A.AppointmentDate,
    V.VaccineName,
    VR.RecordID,
    VR.DoseNo,
    VR.VaccinationDate,
    VC.CenterName,
    B.BatchNumber,
    O.ImmediateReaction,
    O.ObservationOutcome,
    SE.SymptomType,
    SE.SeverityLevel,
    SE.ReportDate,
    SE.FollowUpStatus
FROM (((((((Citizens AS C
LEFT JOIN Appointments AS A ON C.CitizenID = A.CitizenID)
LEFT JOIN VaccinationRecords AS VR ON A.AppointmentID = VR.AppointmentID)
LEFT JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
LEFT JOIN VaccinationCenters AS VC ON A.CenterID = VC.CenterID)
LEFT JOIN VaccineBatches AS B ON VR.BatchID = B.BatchID)
LEFT JOIN ObservationRecords AS O ON VR.RecordID = O.RecordID)
LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
ORDER BY C.CitizenID, VR.VaccinationDate;
"@

    $map['qry_VaccineUsageByBrand'] = @"
SELECT
    V.VaccineName,
    Count(VR.RecordID) AS VaccinationCount
FROM Vaccines AS V
LEFT JOIN VaccinationRecords AS VR ON V.VaccineID = VR.VaccineID
GROUP BY V.VaccineName
HAVING Count(VR.RecordID) > 0
ORDER BY Count(VR.RecordID) DESC, V.VaccineName;
"@

    $map['qry_SideEffectByBrandSeverity'] = @"
SELECT
    V.VaccineName,
    SE.SeverityLevel,
    Count(SE.ReportID) AS ReportCount
FROM ((VaccinationRecords AS VR
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
INNER JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
GROUP BY V.VaccineName, SE.SeverityLevel
ORDER BY V.VaccineName, SE.SeverityLevel;
"@

    $map['qry_DistrictVaccinationRate'] = @"
SELECT
    C.District,
    Count(C.CitizenID) AS TotalCitizens,
    Sum(IIf(Not IsNull(VRC.CitizenID),1,0)) AS VaccinatedCitizens,
    IIf(Count(C.CitizenID)=0,0,Sum(IIf(Not IsNull(VRC.CitizenID),1,0))*100.0/Count(C.CitizenID)) AS VaccinationRatePercent
FROM Citizens AS C
LEFT JOIN (
    SELECT CitizenID
    FROM VaccinationRecords
    GROUP BY CitizenID
)  AS VRC ON C.CitizenID = VRC.CitizenID
GROUP BY C.District
ORDER BY C.District;
"@

    $map['qry_NoShowRate'] = @"
SELECT
    VC.CenterName,
    Count(A.AppointmentID) AS TotalAppointments,
    Sum(IIf([A].[AppointmentStatus]='NoShow',1,0)) AS NoShowCount,
    IIf(Count(A.AppointmentID)=0,0,Sum(IIf([A].[AppointmentStatus]='NoShow',1,0))*100.0/Count(A.AppointmentID)) AS NoShowRatePercent
FROM VaccinationCenters AS VC
LEFT JOIN Appointments AS A ON VC.CenterID = A.CenterID
GROUP BY VC.CenterName
HAVING Count(A.AppointmentID) > 0
ORDER BY 4 DESC, VC.CenterName;
"@

    $map['qry_Top5BusyCenters'] = @"
SELECT TOP 5
    VC.CenterName,
    Count(VR.RecordID) AS TotalVaccinations
FROM (VaccinationCenters AS VC
INNER JOIN Appointments AS A ON VC.CenterID = A.CenterID)
INNER JOIN VaccinationRecords AS VR ON A.AppointmentID = VR.AppointmentID
GROUP BY VC.CenterName
ORDER BY Count(VR.RecordID) DESC, VC.CenterName;
"@

    $map['qry_AgeGroupDistribution'] = @"
SELECT
    X.AgeGroup,
    Count(X.CitizenID) AS CitizenCount
FROM (
    SELECT
        Citizens.CitizenID,
        $ageGroupExpr AS AgeGroup
    FROM Citizens
)  AS X
GROUP BY X.AgeGroup
ORDER BY X.AgeGroup;
"@

    $map['qry_AgeGroupSideEffectRate'] = @"
SELECT
    X.AgeGroup,
    Count(X.RecordID) AS VaccinationRecords,
    Sum(IIf(Not IsNull(X.ReportID),1,0)) AS SideEffectReports,
    IIf(Count(X.RecordID)=0,0,Sum(IIf(Not IsNull(X.ReportID),1,0))*100.0/Count(X.RecordID)) AS SideEffectRatePercent
FROM (
    SELECT
        VR.RecordID,
        $ageGroupExpr AS AgeGroup,
        SE.ReportID
    FROM (VaccinationRecords AS VR
    INNER JOIN Citizens ON VR.CitizenID = Citizens.CitizenID)
    LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID
)  AS X
GROUP BY X.AgeGroup
ORDER BY X.AgeGroup;
"@

    $map['qry_VaccineBrandReactionComparison'] = @"
SELECT
    V.VaccineName,
    Count(VR.RecordID) AS VaccinationRecords,
    Sum(IIf(Not IsNull(SE.ReportID),1,0)) AS SideEffectReports,
    IIf(Count(VR.RecordID)=0,0,Sum(IIf(Not IsNull(SE.ReportID),1,0))*100.0/Count(VR.RecordID)) AS ReactionRatePercent
FROM ((VaccinationRecords AS VR
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
LEFT JOIN SideEffects AS SE ON VR.RecordID = SE.VaccinationRecordID)
GROUP BY V.VaccineName
HAVING Count(VR.RecordID) > 0
ORDER BY 4 DESC, V.VaccineName;
"@

    $map['qry_ScreeningRiskFlags'] = @"
SELECT
    S.ScreeningID,
    C.CitizenName,
    A.AppointmentDate,
    S.Temperature,
    S.RiskFlag,
    S.AllergyConfirmed,
    S.ContraindicationNote
FROM (PreVaccinationScreening AS S
INNER JOIN Citizens AS C ON S.CitizenID = C.CitizenID)
INNER JOIN Appointments AS A ON S.AppointmentID = A.AppointmentID
WHERE S.RiskFlag = True
   OR S.Temperature >= 37.5
   OR S.AllergyConfirmed = True
ORDER BY S.ScreeningDate DESC, C.CitizenName;
"@

    $map['qry_ObservationImmediateReactions'] = @"
SELECT
    O.ObservationID,
    C.CitizenName,
    V.VaccineName,
    VR.DoseNo,
    O.ObservationStart,
    O.ObservationEnd,
    O.ObservationOutcome
FROM (((ObservationRecords AS O
INNER JOIN VaccinationRecords AS VR ON O.RecordID = VR.RecordID)
INNER JOIN Citizens AS C ON O.CitizenID = C.CitizenID)
INNER JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
WHERE O.ImmediateReaction = True
ORDER BY O.ObservationStart DESC, C.CitizenName;
"@

    $map['qry_SlotUtilizationByCenter'] = @"
SELECT
    VC.CenterName,
    Count(S.SlotID) AS TotalSlots,
    Sum(S.BookedCount) AS TotalBookings,
    Sum(S.Capacity) AS TotalCapacity,
    IIf(Sum(S.Capacity)=0,0,Sum(S.BookedCount)*100.0/Sum(S.Capacity)) AS UtilizationPercent
FROM AppointmentSlots AS S
INNER JOIN VaccinationCenters AS VC ON S.CenterID = VC.CenterID
GROUP BY VC.CenterName
HAVING Count(S.SlotID) > 0
ORDER BY 5 DESC, VC.CenterName;
"@

    $map['qry_InventoryUsageAndWastage'] = @"
SELECT
    VC.CenterName,
    V.VaccineName,
    Sum(IIf([T].[TransactionType]='Used',[T].[Quantity],0)) AS UsedQty,
    Sum(IIf([T].[TransactionType]='Wastage',[T].[Quantity],0)) AS WastageQty,
    Sum(IIf([T].[TransactionType]='Adjustment',[T].[Quantity],0)) AS AdjustmentQty
FROM ((InventoryTransactions AS T
INNER JOIN VaccineBatches AS B ON T.BatchID = B.BatchID)
INNER JOIN Vaccines AS V ON B.VaccineID = V.VaccineID)
INNER JOIN VaccinationCenters AS VC ON T.CenterID = VC.CenterID
GROUP BY VC.CenterName, V.VaccineName
HAVING Sum(IIf([T].[TransactionType]='Used',[T].[Quantity],0)) > 0
    OR Sum(IIf([T].[TransactionType]='Wastage',[T].[Quantity],0)) > 0
ORDER BY 4 DESC, 3 DESC, VC.CenterName;
"@

    $map['qry_FollowUpOpenCases'] = @"
SELECT
    F.FollowUpID,
    C.CitizenName,
    V.VaccineName,
    SE.SymptomType,
    SE.SeverityLevel,
    F.FollowUpDate,
    F.ActionTaken,
    F.FollowUpOutcome,
    F.EscalationFlag
FROM ((((AdverseEffectFollowUps AS F
INNER JOIN SideEffects AS SE ON F.ReportID = SE.ReportID)
INNER JOIN Citizens AS C ON SE.CitizenID = C.CitizenID)
LEFT JOIN VaccinationRecords AS VR ON SE.VaccinationRecordID = VR.RecordID)
LEFT JOIN Vaccines AS V ON VR.VaccineID = V.VaccineID)
WHERE F.FollowUpOutcome IN ('Open','Monitoring','Needs Review')
   OR F.EscalationFlag = True
ORDER BY F.EscalationFlag DESC, F.FollowUpDate DESC;
"@

    return $map
}

function Save-SqlArtifacts([string[]]$createSql, [string[]]$indexSql, [System.Collections.Generic.List[string]]$insertSql, [System.Collections.Specialized.OrderedDictionary]$querySqlMap) {
    if (-not (Test-Path $sqlRoot)) {
        New-Item -ItemType Directory -Path $sqlRoot | Out-Null
    }
    if (-not (Test-Path $querySqlRoot)) {
        New-Item -ItemType Directory -Path $querySqlRoot | Out-Null
    }

    Set-Content -Path (Join-Path $sqlRoot '01_create_tables.sql') -Value (($createSql -join "`r`n") + "`r`n")
    Set-Content -Path (Join-Path $sqlRoot '02_relationships_and_indexes.sql') -Value (($indexSql -join "`r`n") + "`r`n")
    Set-Content -Path (Join-Path $sqlRoot '03_insert_sample_data.sql') -Value (($insertSql -join "`r`n") + "`r`n")

    foreach ($file in Get-ChildItem -Path $querySqlRoot -Filter '*.sql' -ErrorAction SilentlyContinue) {
        Remove-Item -Path $file.FullName -Force
    }
    foreach ($queryName in $querySqlMap.Keys) {
        Set-Content -Path (Join-Path $querySqlRoot ($queryName + '.sql')) -Value ($querySqlMap[$queryName] + "`r`n")
    }

    $readme = @'
# Access SQL Deliverables

Database: `gp.accdb`

This v2 package rebuilds the vaccine registration system with:
- 17 tables
- 40 citizens
- 30 vaccines
- 30 vaccination centers
- 40 vaccine batches
- 60 appointment slots
- 40 appointments
- 36 vaccination records
- 40 screenings
- 40 consent records
- 36 observation records
- 35 side effect reports
- 30 adverse-effect follow-ups
- 40 notifications
- 40 vaccine allocation records
- 40 inventory transactions
- 40 audit-log rows

Execution order in Microsoft Access:
1. Run `01_create_tables.sql`
2. Run `02_relationships_and_indexes.sql`
3. Run `03_insert_sample_data.sql`
4. Import or create the saved queries from the `queries` folder
'@
    Set-Content -Path (Join-Path $sqlRoot 'README.md') -Value ($readme + "`r`n")
    Log 'SQL ARTIFACTS WRITTEN'
}

$tableColumns = [ordered]@{
    Citizens = @('HKID','CitizenName','FirstName','LastName','Gender','DOB','Phone','Email','Address','District','AllergyHistory','Occupation','PlaceOfBirth')
    Vaccines = @('VaccineName','Manufacturer','RequiredDoses','DoseIntervalDays','StorageTemp','MinAge','MaxAge')
    VaccinationCenters = @('CenterName','Address','Region','ContactPhone','OperatingHours','MaxDailyService','DoctorName','IsActive')
    VaccineBatches = @('BatchNumber','VaccineID','ManufactureDate','ExpiryDate','CurrentStock','Supplier')
    Users = @('Username','UserPassword','Role','LinkedCitizenID','IsActive')
    AppointmentSlots = @('CenterID','VaccineID','SlotDate','SlotTime','Capacity','BookedCount','SlotStatus')
    Appointments = @('CitizenID','VaccineID','CenterID','SlotID','BatchID','AppointmentDate','DoseNo','AppointmentStatus','ActualVaccinationDate','VaccinatedBy','Notes','CancellationReason')
    VaccinationRecords = @('AppointmentID','CitizenID','VaccineID','BatchID','DoseNo','VaccinationDate','StaffUserID','OutcomeStatus')
    PreVaccinationScreening = @('AppointmentID','CitizenID','ScreeningDate','RiskFlag','Temperature','AllergyConfirmed','ContraindicationNote','ScreenedBy')
    ConsentRecords = @('AppointmentID','CitizenID','ConsentStatus','ConsentDate','WitnessedBy','ConsentNote')
    ObservationRecords = @('RecordID','CitizenID','ObservationStart','ObservationEnd','ImmediateReaction','ObservationOutcome','ObservedBy')
    SideEffects = @('VaccinationRecordID','AppointmentID','CitizenID','SymptomType','SeverityLevel','ReportDate','Details','FollowUpStatus')
    AdverseEffectFollowUps = @('ReportID','FollowUpDate','StaffUserID','ActionTaken','FollowUpOutcome','EscalationFlag')
    Notifications = @('CitizenID','NotificationType','Message','SendDate','IsRead','RelatedAppointmentID')
    VaccineAllocation = @('BatchID','CenterID','QuantityAllocated','AllocationDate','AllocatedBy')
    InventoryTransactions = @('BatchID','CenterID','TransactionType','Quantity','TransactionDate','RelatedRecordID','ProcessedBy')
    AuditLog = @('UserID','Action','TableName','RecordID','ActionDate','Details')
}

try {
    if (Test-Path $backupPath) {
        Remove-Item -Path $backupPath -Force
    }
    Copy-Item -Path $databasePath -Destination $backupPath -Force
    Log 'BACKUP CREATED'

    $createSql = Build-CreateSql
    $indexSql = Build-IndexSql
    $insertSql = [System.Collections.Generic.List[string]]::new()
    $querySqlMap = [ordered]@{}

    $citizenRows = Get-CitizenRows
    $vaccineRows = Get-VaccineRows
    $centerRows = Get-CenterRows
    $batchRows = Get-BatchRows $vaccineRows
    $userRows = Get-UserRows
    $slotRows = Get-AppointmentSlotRows
    $appointmentRows = Get-AppointmentRows $slotRows
    Update-SlotUsage $slotRows $appointmentRows
    $vaccinationRecordRows = Get-VaccinationRecordRows $appointmentRows
    $screeningRows = Get-ScreeningRows $appointmentRows $citizenRows
    $consentRows = Get-ConsentRows $appointmentRows
    $observationRows = Get-ObservationRows $vaccinationRecordRows
    $sideEffectRows = Get-SideEffectRows $vaccinationRecordRows
    $followUpRows = Get-FollowUpRows $sideEffectRows
    $notificationRows = Get-NotificationRows $appointmentRows
    $allocationRows = Get-AllocationRows
    $inventoryRows = Get-InventoryTransactionRows $vaccinationRecordRows
    $auditRows = Get-AuditRows

    Log 'OPEN DAO'
    $dbEngine = New-Object -ComObject DAO.DBEngine.120
    $db = $dbEngine.OpenDatabase($databasePath)
    Log 'DB OPENED'

    foreach ($queryName in @(
        'qry_UI_CitizenProfile','qry_UI_Appointments',
        'qry_AppointmentConflictCheck','qry_SecondDoseReminder','qry_ExpiringOrLowStockBatches','qry_CenterCapacityUsage',
        'qry_CitizenVaccinationHistory','qry_VaccineUsageByBrand','qry_SideEffectByBrandSeverity','qry_DistrictVaccinationRate',
        'qry_NoShowRate','qry_Top5BusyCenters','qry_AgeGroupDistribution','qry_AgeGroupSideEffectRate',
        'qry_VaccineBrandReactionComparison','qry_ScreeningRiskFlags','qry_ObservationImmediateReactions',
        'qry_SlotUtilizationByCenter','qry_InventoryUsageAndWastage','qry_FollowUpOpenCases'
    )) {
        Remove-QueryIfExists $db $queryName
    }

    foreach ($tableName in @(
        'AuditLog','InventoryTransactions','VaccineAllocation','Notifications','AdverseEffectFollowUps','SideEffects',
        'ObservationRecords','ConsentRecords','PreVaccinationScreening','VaccinationRecords','Appointments',
        'AppointmentSlots','Users','VaccineBatches','VaccinationCenters','Vaccines','Citizens'
    )) {
        Remove-TableIfExists $db $tableName
    }

    foreach ($sql in $createSql) {
        $db.Execute($sql, $dbFailOnError)
    }
    foreach ($sql in $indexSql) {
        $db.Execute($sql, $dbFailOnError)
    }
    Log 'SCHEMA READY'

    foreach ($row in $citizenRows) { Execute-Insert $db 'Citizens' $tableColumns.Citizens $row $insertSql }
    foreach ($row in $vaccineRows) { Execute-Insert $db 'Vaccines' $tableColumns.Vaccines $row $insertSql }
    foreach ($row in $centerRows) { Execute-Insert $db 'VaccinationCenters' $tableColumns.VaccinationCenters $row $insertSql }
    foreach ($row in $batchRows) { Execute-Insert $db 'VaccineBatches' $tableColumns.VaccineBatches $row $insertSql }
    foreach ($row in $userRows) { Execute-Insert $db 'Users' $tableColumns.Users $row $insertSql }
    foreach ($row in $slotRows) { Execute-Insert $db 'AppointmentSlots' $tableColumns.AppointmentSlots $row $insertSql }
    foreach ($row in $appointmentRows) { Execute-Insert $db 'Appointments' $tableColumns.Appointments $row $insertSql }
    foreach ($row in $vaccinationRecordRows) { Execute-Insert $db 'VaccinationRecords' $tableColumns.VaccinationRecords $row $insertSql }
    foreach ($row in $screeningRows) { Execute-Insert $db 'PreVaccinationScreening' $tableColumns.PreVaccinationScreening $row $insertSql }
    foreach ($row in $consentRows) { Execute-Insert $db 'ConsentRecords' $tableColumns.ConsentRecords $row $insertSql }
    foreach ($row in $observationRows) { Execute-Insert $db 'ObservationRecords' $tableColumns.ObservationRecords $row $insertSql }
    foreach ($row in $sideEffectRows) { Execute-Insert $db 'SideEffects' $tableColumns.SideEffects $row $insertSql }
    foreach ($row in $followUpRows) { Execute-Insert $db 'AdverseEffectFollowUps' $tableColumns.AdverseEffectFollowUps $row $insertSql }
    foreach ($row in $notificationRows) { Execute-Insert $db 'Notifications' $tableColumns.Notifications $row $insertSql }
    foreach ($row in $allocationRows) { Execute-Insert $db 'VaccineAllocation' $tableColumns.VaccineAllocation $row $insertSql }
    foreach ($row in $inventoryRows) { Execute-Insert $db 'InventoryTransactions' $tableColumns.InventoryTransactions $row $insertSql }
    foreach ($row in $auditRows) { Execute-Insert $db 'AuditLog' $tableColumns.AuditLog $row $insertSql }
    Log 'DATA INSERTED'

    foreach ($entry in (Build-QuerySqlMap).GetEnumerator()) {
        Add-OrReplaceQuery $db $entry.Key $entry.Value $querySqlMap
    }

    Save-SqlArtifacts $createSql $indexSql $insertSql $querySqlMap
    Log 'V2 DATA UPGRADE COMPLETE'
} catch {
    Log ('ERROR: ' + $_.Exception.Message)
    throw
} finally {
    try { if ($db) { $db.Close() } } catch {}
    $db = $null
    $dbEngine = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Log 'END'
}
