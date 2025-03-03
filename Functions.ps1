# Functions.ps1
<#
    Script Name: Functions.ps1
    Purpose    : Helper functions for AD_SweeperV2.ps1 and New_User_Creation.ps1
    Author     : Isidro Paniagua and Devin Lacey
    Date       : 01/17/2025

    Description:
    This script contains utility functions used by AD_SweeperV2.ps1 and New_User_Creation.ps1 for:
    - Converting and standardizing department names
    - Mapping job titles to standard formats
    - Managing Active Directory Organizational Units (OUs)
    - Sending email notifications
    - Text formatting and case conversion

    Dependencies:
    - Active Directory PowerShell module
    - Exchange Server access for email notifications
    - Appropriate AD permissions for OU management
#>

<#
.SYNOPSIS
    Converts department names to standardized format.

.DESCRIPTION
    Takes a department name as input and returns a standardized version based on predefined mappings.
    If no mapping exists, converts the name to proper title case.

.PARAMETER department
    The department name to standardize.

.EXAMPLE
    ConvertDeptName -department "ACCOUNTING FINANCE"
    Returns: "Finance"
#>
function ConvertDeptName {
    param (
        [string]$department
    )
        $department = $department.Trim().ToUpper()
    
    $DepartmentMapping = @{
        'ACCOUNTING FINANCE' = "Finance"
        'ADVERTISING AND MARKETING' = "Marketing"
        'PROMOTIONS  ENTERTAINMENT' = "Marketing"
        'GROUP SALES' = "Marketing"
        'AUDIO VISUAL' = "Audio Visual"
        'BEVERAGE' = "Beverage"
        'BUSINESS INTELLIGENCE' = "Business Intelligence"
        'CAGE' = "Cage"
        'CASINO BANQUETS' = "Casino Banquets"
        'CASINO GIFT SHOP' = "Gift Shop"
        'COUNT' = "Count"
        'EMERALD' = "Emerald"
        'EVS HOUSEKEEPING' = "EVS"
        'TDR' = "TDR"
        'EXECUTIVE' = "Executive"
        'FACILITY ENGINEERING' = "Facilities"
        'FB ADMIN' = "F&B Admin"
        'HIGH LIMIT BAR' = "High Limit Bar"
        'HOTEL ADMIN' = "Hotel Admin"
        'HUMAN RESOURCES' = "Human Resources"
        'LOFT' = "Loft"
        'MARKETPLACE' = "Marketplace"
        'PD VIP SERVICES' = "Player Development"
        'POKER ROOM' = "Poker Room"
        'PURCHASING' = "Purchasing"
        'SECURITY' = "Security"
        'SLOTS' = "Slots"
        'STEWARDING' = "Stewarding"
        'SURVEILLANCE' = "Surveillance"
        'TABLE GAMES' = "Table Games"
        'TONY GWYNNS' = "Tony Gwynns"
        'VALET' = "Valet"
        'WARDROBE' = "Wardrobe"
        'WAREHOUSE RECEIVING' = "Warehouse"
        "PRIME CUT" = "Prime Cut"
    }

    if ($DepartmentMapping.ContainsKey($department)) {
        return $DepartmentMapping[$department]
    }
    return (ConvertToTitleCase $department)
}

<#
.SYNOPSIS
    Maps job titles to standardized formats.

.DESCRIPTION
    Converts various forms of job titles to their standardized versions using a comprehensive mapping dictionary.
    Handles special cases for different departments and ensures consistent naming across the organization.

.PARAMETER adTitle
    The job title to standardize.

.EXAMPLE
    MapJobTitles -adTitle "SLOT TECH SHIFT MANAGER"
    Returns: "Slot Technician Shift Manager"
#>
function MapJobTitles {
    param (
        [string]$adTitle
    )
    
    # Clean up the input by removing extra spaces and converting to uppercase for matching
    $upperTitle = $adTitle.Trim().ToUpper()
    
    $JobTitleMapping = @{
        # Slots
        "SLOT TECH SHIFT MANAGER" = "Slot Technician Shift Manager"
        "SLOT TECH ASST SHIFT MGR" = "Assistant Slot Technician Shift Manager"
        "DIRECTOR OF SLOT OPS" = "Director of Slots"
        "ASSISTANT SLOT SHFT MGR" = "Assistant Slot Shift Manager"
        "SLOT TECHNICAL SUPERVISOR" = "Slot Technician Supervisor"
        
        # Facilities 
        "DIRECTOR OF PROPERTY MGMT" = "Director of Property Management"
        "PROPERTY OPS ADMIN" = "Facilities and EVS Admin Coordinator"

        # EVS
        "EVS SUPERVISOR" = "EVS Supervisor"
        "EVS LEAD" = "EVS Lead"

        # Business Intelligence
        "DIR BI DATA ANALYTICS" = "Director of Business Intelligence"
        "BUSINESS INT ANALYST" = "Business Intelligence Analyst"

        # Compliance
        "SR COMPLIANCE RISK SPEC" = "Senior Compliance Risk Specialist"
        "DIR COMPLIANCE RISK" = "Director of Compliance & Risk"
        
        # Finance
        "AP COORDINATOR" = "A/P Coordinator"
        "ACCOUNTS PAYABLE MANAGER" = "A/P Manager"
        
        # Marketing
        "PROMO ENTRTNMNT MANAGER" = "Promotions and Entertainment Manager"
        "ASST MGR PROMO ENTERTN" = "Assistant Manager Promotions and Entertainment"
        "SR PROMO AND EVENTS COORD" = "Senior Promotions and Events Coordinator"
        "PROMO AND EVENTS COORD" = "Promotions and Events Coordinator"
        "ASST ADVERTISING MANAGER" = "Assistant Advertising Manager"
        "MKTG DATABASE MGR" = "Marketing Database Manager"
        "SOCIAL MEDIA CONTENT SP" = "Social Media and Content Specialist"
        "VP OF MARKETING" = "VP of Marketing"
        "SR CREATIVE PRODUCER" = "Senior Creative Producer"
        "SR GRAPHIC DESIGNER" = "Senior Graphic Designer"
        "SR SPECIAL EVENTS PLANNER" = "Senior Special Events Planner"
        
        # Audio Visual
        "AUDIO VISUAL ASST MANAGER" = "Assistant Audio Visual Manager"
        "SENIOR AV TECH" = "Senior Audio Visual Technician"
        
        # F&B
        "BEVERAGE CART ATTND NA" = "Beverage Cart Attendant"
        "FB RETAIL SUPERVISOR" = "F&B Retail Supervisor"
        "FB OFFICE RETAIL MGR" = "F&B Office Retail Manager"
        "DIRECTOR OF FOOD AND BEV" = "Director of F&B"
        "FB COORDINATOR" = "F&B Coordinator"
        "ASSISTANT MANAGER FB" = "Assistant F&B Manager"
        "FB ADMIN" = "F&B Admin"
        "EVS - STEWARD OPS MANAGER" = "EVS Steward Operations Manager"

        # Guest Services
        "DIRECTOR OF GUEST SVCS" = "Director of Guest Services"
        "GS SUPPORT CASHIER" = "Guest Services Support Cashier"
        "GUEST SERVICE CASHIER" = "Guest Services Cashier"
        "GUEST SERVICE MANAGER" = "Guest Services Manager"
        "GUEST SERVICES SUPERVISOR" = "Guest Services Supervisor"
        "GUEST SVCS SHIFT MGR" = "Guest Services Shift Manager"
        "GS LEAD CASHIER" = "Guest Services Lead Cashier"

        # Emeralds
        "ASIAN RESTAURANT COOK 2" = "Emerald Cook 2"
        "ASIAN REST BUS PERSON" = "Emerald Bus Person"
        "EMERALD HOST-CASHIER" = "Emerald Host - Cashier"
        "FB ASIAN REST SUPERVISOR" = "Emerald Supervisor"
        "ASIAN REST ATTNDNT - LEAD" = "Emerald Lead Attendant"
        "ASIAN REST COOK LEAD" = "Emerald Lead Cook"
        "ASIAN REST EXPEDITOR" = "Emerald Expeditor"

        # Executive
        "PRESIDENT - GM" = "President / General Manager"
        "DIR STRATEGIC PLANNING" = "Director of Strategic Planning"
        "VP NON-GAMING" = "VP of Non-Gaming Operations"
        "VP OF CASINO OPERATIONS" = "VP of Casino Operations"
        "MANAGER ON DUTY" = "Manager on Duty"

        # Player Development
        "VIP SERVICES MANAGER" = "VIP Services Manager"
        "VIP SERVICES ADMIN" = "VIP Services Admin"
        "VIP SERVICES SPECIALIST" = "VIP Services Specialist"
        "DIR PLAYER DEV VIP SVCS" = "Director of Player Development VIP Services"
        "SR EXECUTIVE CASINO HOST" = "Senior Executive Casino Host"

        # Guest Relations
        "GUEST RELATIONS SPEC" = "Guest Relations Specialist"
        
        # Hotel
        "DIRECTOR OF HOTEL" = "Director of Hotel"

        # Human Resources
        "HR CONCIERGE" = "HR Concierge"
        "HR BENEFITS ADMINISTRATOR" = "HR Benefits Administrator"
        "ASSISTANT HR MANAGER" = "Assistant HR Manager"
        "SR HR GENERALIST-HRIS" = "Senior HR Generalist-HRIS"
        "HUMAN RESOURCES COORDIN" = "HR Coordinator"
        "HR SPECIALIST- TRAINING" = "HR Training Specialist"
        "HR MANAGER" = "HR Manager"

        # Loft
        "BEER GARDEN LEAD COOK" = "Loft Lead Cook"
        "BEER GARDEN SERVER" = "Loft Server"
        "BEER GARDEN HOST" = "Loft Host"
        "BEER GARDEN BARTENDER" = "Loft Bartender"
        "ASSISTANT F&B MANAGER" = "Assistant F&B Manager"
        "BEER GARDEN BUS PERSON" = "Loft Bus Person"

        # Marketplace
        "FOOD COURT ATTENDANT" = "Marketplace Attendant"
        "FOOD COURT BUS PERSON" = "Marketplace Bus Person"
        "FB FOOD COURT SPVSR" = "Marketplace Supervisor"
        "MARKETPLACE SPVSR" = "Marketplace Supervisor"
        "FOOD COURT COOK 1" = "Marketplace Cook 1"
        "FOOD COURT COOK 2" = "Marketplace Cook 2"
        "FB FOOD COURT MANAGER" = "Marketplace Manager"
        "FOOD COURT ATTND - LEAD" = "Marketplace Lead Attendant"
        "FOOD COURT LEAD COOK" = "Marketplace Lead Cook"

        # Poker Room
        "POKER DEALER" = "Poker Dealer Dual Rate Supervisor"

        # Prime Cut
        "STEAKHOUSE HOST" = "Prime Cut Host"
        "STEAKHOUSE BARBACK" = "Prime Cut Barback"
        "STEAKHOUSE SERVER" = "Prime Cut Server"
        "STEAKHOUSE ASSIST SERVER" = "Assistant Prime Cut Server"
        "STEAKHOUSE BARTENDER" = "Prime Cut Bartender"
        "FB STEAKHOUSE MANAGER" = "Prime Cut Manager"

        # Jive
        "LOUNGE BEVERAGE SERVER" = "Jive Beverage Server"
        "LOUNGE BARTENDER" = "Jive Bartender"
        "LOUNGE HOST" = "Jive Host"

        # Tony Gwynns
        "SPORTS BAR BARTENDER" = "Tony Gwynns Bartender"
        "SPORTS BAR BUS PERSON" = "Tony Gwynns Bus Person"
        "SPORTS BAR HOST - CASHIER" = "Tony Gwynns Host - Cashier"
        "SPORTS BAR SERVER" = "Tony Gwynns Server"
        "FB SPORTS BAR MANAGER" = "Tony Gwynns Manager"
        "FB SPORTS BAR SUPERVISOR" = "Tony Gwynns Supervisor"

        # Purchasing
        "SR BUYER" = "Senior Buyer"
        "DIRECTOR OF PURCHASING" = "Director of Purchasing"

        # Security
        "DIRECTOR OF SECURITY" = "Director of Security"
        "SECURITY OPS TRAINING MGR" = "Security Ops Training Manager"
        "SECURITY SHIFT MGR" = "Security Shift Manager"
        "SECURITY SUPERVISOR- GRV" = "Security Supervisor - Grave"
        "SECURITY OFFICER- GRV" = "Security Officer - Grave"
        "SECURITY COMMS MANAGER" = "Security Communications Manager"
        "SECURITY DISPATCHER GRV" = "Security Dispatcher - Grave"
        "EXT PATROL OFFICER GRV" = "Exterior Patrol Officer - Grave"
        "SECURITY OFFICER-EMT GRV" = "Security Officer - EMT Grave"
        "SECURITY OFFICER LEAD-GRV" = "Lead Security Officer - Grave"
        "SECURITY OFFICER-LEAD" = "Lead Security Officer"
        "SECURITY OFFICER - EMT" = "Security Officer - EMT"

        # Surveillance
        "DIRECTOR OF SURVEILLANCE" = "Director of Surveillance"
        "SURVEILLANCE SHFT MGR" = "Surveillance Shift Manager"
        "SURVEILLANCE LEAD AGENT" = "Lead Surveillance Agent"
        "SURVEILLANCE AGENT" = "Surveillance Agent"

        # Table Games
        "ASST TABLDE GAMES SHFT MGR" = "Assistant Table Games Shift Manager"
        "DIRECTOR OF TABLE GAMES" = "Director of Table Games"
        "DEALER" = "Poker Dealer Dual Rate Supervisor"
    }

    if ($JobTitleMapping.ContainsKey($upperTitle)) {
        return $JobTitleMapping[$upperTitle]
    } else {
        return (ConvertToTitleCase $adTitle)
    }
}

<#
.SYNOPSIS
    Sends email notifications about script actions.

.DESCRIPTION
    Sends automated email notifications using the organization's Exchange server.
    Includes audit information about who ran the script and what action was taken.

.PARAMETER message
    The message content to send.
.PARAMETER recipient
    The email address of the recipient.

.EXAMPLE
    sendEmail -message "New user account created" -recipient "admin@domain.com"
#>
function sendEmail {
    param (
    $message,
    $recipient

    )
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $smtpServer = "jcexchange2019.jamulcasinosd.com"
    $from = "ScriptAlert@jamulcasinosd.com"
    $subject = "User Creation Alert"


    # Prepare the email body
    $body = "$message
    
            Script ran by: $currentUser"

    # Send the email
    Send-MailMessage -SmtpServer $smtpServer -From $from -To $recipient -Subject $subject -Body $body
}

<#
.SYNOPSIS
    Converts strings to proper title case.

.DESCRIPTION
    Converts input strings to proper title case format, removing apostrophes and handling capitalization.
    Uses .NET TextInfo for proper word capitalization.

.PARAMETER myString
    The string to convert to title case.

.EXAMPLE
    ConvertToTitleCase -myString "HUMAN RESOURCES"
    Returns: "Human Resources"
#>
function ConvertToTitleCase {
    param (
        [string]$myString
    )

    # Remove apostrophes from the string
    $myString = $myString -replace "'", ""

    $textInfo = (Get-Culture).TextInfo
    $titleCaseString = $textInfo.ToTitleCase($myString.ToLower())

    return $titleCaseString
}

# Define the OU Mapping using a hash table for faster lookups
$OUMapping = @{
    "Guest Services Cashier"            = "OU=Guest Services Cashier,OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Guest Services Support Cashier"    = "OU=Guest Services Cashier,OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "VIP Services Manager"              = "OU=VIP Services Manager,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Security"              = "OU=Security Director,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Chief Financial Officer"           = "OU=Exec,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Executive Assistant"               = "OU=Exec,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Manager on Duty"                   = "OU=Exec,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "President / General Manager"       = "OU=Exec,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Manager"                  = "OU=Security Manager,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Ops Training Manager"     = "OU=Security Manager,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Shift Manager"            = "OU=Security Manager,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Pit Manager"             = "OU=Assistant Pit Manager,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Table Games Shift Manager" = "OU=Assistant Pit Manager,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Guest Services Shift Manager"      = "OU=Guest Services Shift Manager,OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Guest Services Manager"            = "OU=Guest Services Shift Manager,OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Technician Shift Manager"     = "OU=Slot Technician Shift Manager,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Attendant"                    = "OU=Slot Attendant,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Casino Concierge Temp"             = "OU=Slot Attendant,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Count Room Manager"                = "OU=Count Room Manager,OU=Count,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Buyer"                             = "OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Purchasing"            = "OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Buyer"                      = "OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Surveillance"          = "OU=Director of Surveillance and Transportation,OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Surveillance Shift Manager"        = "OU=Surveillance Shift Manager,OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Guest Services Supervisor"         = "OU=Guest Services Supervisor,OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Technician Supervisor"        = "OU=Slot Technical Supervisor,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Audio Visual Manager"    = "OU=AV Managers,OU=Audio Visual,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Audio Visual Manager"              = "OU=AV Managers,OU=Audio Visual,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Advertising Manager"     = "OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Group Sales Specialist"            = "OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Facilities Lead"                   = "OU=Facilities Lead,OU=Facilities,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Facilities Supervisor"             = "OU=Facilities Lead,OU=Facilities,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Revenue Auditor"                   = "OU=Revenue Auditor,OU=Finance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Group Sales Manager"               = "OU=Group Sales Manager,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Table Games Supervisor"            = "OU=Table Games Supervisor,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Creative Producer"          = "OU=Senior Graphic Artist,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Graphic Designer"           = "OU=Senior Graphic Artist,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Facilities General Repair"         = "OU=Facilities General Maintenance,OU=Facilities,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Guest Relations Specialist"        = "OU=Guest Relations,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Dispatcher"               = "OU=Security Dispatcher,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Dispatcher - Grave"       = "OU=Security Dispatcher,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Hotel"                 = "OU=Hotel Director,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Hotel Manager"                     = "OU=Hotel Manager,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Executive Housekeeper"             = "OU=Hotel Housekeeping,OU=EVS,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Housekeeping Supervisor"           = "OU=Hotel Housekeeping,OU=EVS,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant HR Manager"              = "OU=Sr. HR Specialist,OU=Human Resources,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Graphic Artist"             = "OU=Senior Graphic Artist,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Manager Promotions and Entertainment" = "OU=Promotions and Entertainment,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Marketing Analyst"                 = "OU=Marketing Analyst,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Marketing Coordinator"             = "OU=Marketing Coordinator,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Promotions and Events Coordinator" = "OU=Promotions and Entertainment,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Promotions and Events Coordinator" = "OU=Promotions and Entertainment,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Marketing Ambassador"              = "OU=Marketing Ambassador,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Community Relations Manager"       = "OU=Community Relations Manager,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Special Events Planner"     = "OU=Special Events Planner,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director Of Advertising"           = "OU=Advertising Manager,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Advertising Manager"               = "OU=Advertising Manager,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Marketing Database Manager"        = "OU=Database Marketing Manager,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Advertising Specialist"            = "OU=Advertising Specialist,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Social Media and Content Specialist" = "OU=Social Media and Content Specialist,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director Of Marketing"             = "OU=Director Of Marketing,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Player Development VIP Services" = "OU=Player Development Manager,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Executive Casino Host"      = "OU=Senior Executive Casino Host,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "VIP Services Admin"                = "OU=VIP Services Specialist,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "VIP Services Specialist"           = "OU=VIP Services Specialist,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Executive Casino Host"             = "OU=Executive Casino Host,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Casino Host"                       = "OU=Casino Host,OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Poker Manager"                     = "OU=Poker Manager,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Poker Supervisor"                  = "OU=Poker Supervisor,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Poker Room Host"                   = "OU=Poker Room Host,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Poker Dealer Dual Rate Supervisor" = "OU=Poker Dealer Dual Rate Supervisor,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Supervisor"               = "OU=Security Supervisor,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Supervisor - Grave"       = "OU=Security Supervisor,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Officer"                  = "OU=Security Officer II,OU=Security Officer,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Coordinator"              = "OU=Security Coordinator,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Administrator"            = "OU=Security Coordinator,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Communications Manager"   = "OU=Security Comms Supv,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Exterior Patrol Officer"           = "OU=Exterior Patrol Officer,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Exterior Patrol Officer - Grave"   = "OU=Exterior Patrol Officer,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Lead Security Officer"             = "OU=Security Leads,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Lead Security Officer - Grave"     = "OU=Security Leads,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Officer - Grave"          = "OU=Security Officer - Grave,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Officer - EMT"            = "OU=Security Officer / EMT,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Security Officer - EMT Grave"      = "OU=Security Officer / EMT,OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Supervisor"                   = "OU=Slot Supervisors,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Slots"                 = "OU=Director of Slots,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Technician Supervisor - Grave" = "OU=Slot Technical Supervisor,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Slot Shift Manager"      = "OU=Assistant Slots Shift Manager,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Shift Manager"                = "OU=Assistant Slots Shift Manager,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant Slot Technician Shift Manager" = "OU=Slot Technician Assistant Shift Manager,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Slot Technician"                   = "OU=Slot Technician,OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Spa Manager"                       = "OU=Spa,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Executive Steward"                 = "OU=Stewarding,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Lead Surveillance Agent"           = "OU=Surveillance Lead Agent,OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Surveillance Agent"                = "OU=Surveillance Agents,OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Surveillance Technician"           = "OU=Surveillance Technician,OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Table Games"           = "OU=Director of Table Games,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Table Games Shift Manager"         = "OU=Assistant Table Games Shift Manager,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Front Services Manager"            = "OU=Transportation Manager,OU=Valet,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Transportation Supervisor"         = "OU=Transportation Supervisor,OU=Valet,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Transportation Manager"            = "OU=Transportation Manager,OU=Valet,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Wardrobe Attendant"                = "OU=Wardrobe,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Wardrobe Supervisor"               = "OU=Wardrobe,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Warehouse Manager"                 = "OU=Warehouse Manager,OU=Warehouse,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Warehouse Lead Person"             = "OU=Warehouse Person,OU=Warehouse,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Warehouse Person"                  = "OU=Warehouse Person,OU=Warehouse,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Warehouse Supervisor"              = "OU=Warehouse Supervisor,OU=Warehouse,OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Beverage Server"                   = "OU=Beverage,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Beverage Cart Attendant"           = "OU=Beverage Cart Attendant,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Assistant F&B Manager"             = "OU=Supervisors,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Beverage Supervisor"               = "OU=Supervisors,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Stewarding Supervisor"             = "OU=Stewarding,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Barback"                           = "OU=Barback,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Prime Cut Barback"                 = "OU=Barback,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Financial Accountant"              = "OU=Financial Accountant,OU=Finance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Marketing Ambassador Lead"         = "OU=Marketing Ambassador Lead,OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Audio Visual Technician"           = "OU=AV Techs,OU=Audio Visual,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Senior Audio Visual Technician"    = "OU=AV Techs,OU=Audio Visual,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    "Director of Business Intelligence" = "OU=Director of BI Data Analytics,OU=Business Intelligence,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
}

<#
.SYNOPSIS
    Maps departments to their corresponding Organizational Units.

.DESCRIPTION
    Takes a department name and returns the corresponding OU path in Active Directory.
    Includes fallback to default OU if no specific mapping is found.

.PARAMETER department
    The department name to map to an OU.

.EXAMPLE
    DepartmentMapping -department "Marketing"
    Returns: "OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
#>
function DepartmentMapping {
    param (
        [string]$department
    )
    Write-Debug "Mapping department to OU: $department"

    # Define department to OU mapping
    $departmentOUMapping = @{
        "Audio Visual" = "OU=Audio Visual,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Business Intelligence" = "OU=Business Intelligence,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Casino Banquets" = "OU=Casino Banquets,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Compliance" = "OU=Compliance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Count" = "OU=Count,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "EVS" = "OU=EVS,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Executive" = "OU=Exec,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Facilities" = "OU=Facilities,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Finance" = "OU=Finance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Food and Beverage" = "OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "F&B Admin" = "OU=FB Admin,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "F&B" = "OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "TDR" = "OU=TDR,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Emeralds" = "OU=Emerald,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Prime Cut" = "OU=Prime Cut,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Jamul 23 Restaurant Bar" = "OU=Jamul 23 Restaurant Bar,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Tony Gwynns" = "OU=Tony Gwynns,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Jive" = "OU=JIVe,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Gift Shop" = "OU=Retail Gift Shop,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "High Limit Bar" = "OU=High Limit Bar,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Loft" = "OU=Loft,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Beverage" = "OU=Beverage,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Stewarding" = "OU=Stewarding,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Marketplace" = "OU=Marketplace,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Hotel Admin" = "OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Spa" = "OU=Spa,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Front Desk" = "OU=Front Desk,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Hotel Housekeeping" = "OU=Hotel Housekeeping,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Starlite Pool" = "OU=Starlite Pool,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Guest Relations" = "OU=Guest Relations,OU=Hotel Admin,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Human Resources" = "OU=Human Resources,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Marketing" = "OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Player Development" = "OU=Player Development,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Player Services" = "OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Cage" = "OU=Player Services,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Purchasing" = "OU=Purchasing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Security" = "OU=Security,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Slots" = "OU=Slots,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Surveillance" = "OU=Surveillance,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Table Games" = "OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Poker Room" = "OU=Table Games,OU=Table Games,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Valet" = "OU=Valet,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    }

    # Return the mapped OU or default if not found
    $mappedOU = $departmentOUMapping[$department]
    if ($null -eq $mappedOU) {
        $mappedOU = "OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    }
    Write-Debug "Mapped department '$department' to OU: $mappedOU"
    return $mappedOU
}

<#
.SYNOPSIS
    Maps job titles to appropriate OUs with venue-specific handling.

.DESCRIPTION
    Determines the correct OU for a user based on their job title and department.
    Handles special cases for venue-specific roles and provides fallback options.

.PARAMETER jobTitle
    The job title to map to an OU.
.PARAMETER department
    The department name to use as fallback for OU mapping.

.EXAMPLE
    MapOU -jobTitle "Emerald Host" -department "Food and Beverage"
    Returns: "OU=Emerald,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
#>
function MapOU {
    param (
        [string]$jobTitle,
        [string]$department
    )

    # Venue-specific OU mappings
    $venueOUMapping = @{
        "Emeralds"    = "OU=Emerald,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Tony Gwynns" = "OU=Tony Gwynns,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Jive"        = "OU=JIVe,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Loft"        = "OU=Loft,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Marketplace" = "OU=Marketplace,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Prime Cut"   = "OU=Prime Cut,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Jamul 23 Restaurant Bar" = "OU=Jamul 23,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
        "Gift Shop" = "OU=Retail Gift Shop,OU=Food and Beverage,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
    }

    # Check for venue-specific titles
    foreach ($key in $venueOUMapping.Keys) {
        if ($jobTitle -match $key) {
            Write-Debug "Matched job title '$jobTitle' to venue-specific OU: $($venueOUMapping[$key])"
            return $venueOUMapping[$key]
        }
    }

    # Default mapping using the existing OUMapping
    if ($OUMapping.ContainsKey($jobTitle)) {
        Write-Debug "Mapped job title '$jobTitle' to OU: $($OUMapping[$jobTitle])"
        return $OUMapping[$jobTitle]
    } else {
        # Call DepartmentMapping if no job title match is found
        $defaultOU = DepartmentMapping -department $department
        Write-Debug "Defaulted to department mapping for '$department': $defaultOU"
        return $defaultOU
    }
}

<#
.SYNOPSIS
    Moves an AD user to the specified Organizational Unit.

.DESCRIPTION
    Attempts to move an Active Directory user object to a new OU location.
    Includes error handling and logging for failed operations.

.PARAMETER AdUser
    The AD user object to move.
.PARAMETER correctOU
    The target OU path where the user should be moved.

.EXAMPLE
    UpdateUserOU -AdUser $userObject -correctOU "OU=Marketing,OU=Users,OU=Jamul,DC=jamulcasinosd,DC=com"
#>
function UpdateUserOU {
    param (
        [PSCustomObject]$AdUser,
        [string]$correctOU
    )

    try {
        # Move the user to the correct OU
        Move-ADObject -Identity $AdUser.DistinguishedName -TargetPath $correctOU
        # Write-Host "Moved user '$($AdUser.Name)' to '$correctOU'." -ForegroundColor Green
    }
    catch {
        Write-Host "Failed to move user '$($AdUser.Name)' to '$correctOU': $($_.Exception.Message)" -ForegroundColor Red
    }
}


