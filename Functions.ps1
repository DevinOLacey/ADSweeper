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

        # Table Games
        "ASST TABLDE GAMES SHFT MGR" = "Assistant Table Games Shift Manager"
        "DIRECTOR OF TABLE GAMES" = "Director of Table Games"
        "DEALER" = "Dual Rate Supervisor and Dealer"
    }

    if ($JobTitleMapping.ContainsKey($upperTitle)) {
        return $JobTitleMapping[$upperTitle]
    } else {
        return (ConvertToTitleCase $adTitle)
    }
}

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


