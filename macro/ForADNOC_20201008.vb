Private Sub CommandButton1_Click()
Dim SourceRange3 As Range, cel As Range

On Error Resume Next

   Set SourceRange3 = Application.Selection
   Set SourceRange3 = Application.InputBox("Range:", "Selece Filenames: ", SourceRange3.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0
    
    SourceRange3.Offset(0, 47).Value = "=UPPER(RC[-47])"
        
    For Each cel In SourceRange3.Offset(0, 47)
    
        If InStr(1, cel.Value, "UTILITY FLOW DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Utility Flow Diagram"
            
        ElseIf InStr(1, cel.Value, "PROCESS FLOW DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Flow Diagram"
          
        ElseIf InStr(1, cel.Value, "AUXILIARY FLOW DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Flow Diagram"
                
        ElseIf InStr(1, cel.Value, "P&ID") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "P&ID"
            
        ElseIf InStr(1, cel.Value, "P & ID") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "P&ID"
                            
        ElseIf InStr(1, cel.Value, "LOOP") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Loop Drawings"
            
        ElseIf InStr(1, cel.Value, "ISOMETRIC") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Fabrication Isometric Drawing"
            
        ElseIf InStr(1, cel.Value, "CABLE BLOCK DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Cable Block Diagram"
            
        ElseIf InStr(1, cel.Value, "PIPELINE ALIGNMENT") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline Alignment Sheets"
            
        ElseIf InStr(1, cel.Value, "ALIGNMENT SHEET") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline Alignment Sheets"
            
       ElseIf InStr(1, cel.Value, "COMPLETION STAGE ARCHITECTURAL DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Architectural"
            cel.Offset(0, -38).Value = "Completion Stage Architectural Drawings"

       ElseIf InStr(1, cel.Value, "COMPLETION STAGE ARCHITECTURAL SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Architectural"
            cel.Offset(0, -38).Value = "Completion Stage Architectural Specification"

       ElseIf InStr(1, cel.Value, "GENERAL ARRANGEMENT FOR BUILDINGS AND STRUCTURES") > 0 Then
            cel.Offset(0, -39).Value = "Architectural"
            cel.Offset(0, -38).Value = "General Arrangement for Buildings and Structures"

       ElseIf InStr(1, cel.Value, "INTERMEDIATE STAGE ARCHITECTURAL DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Architectural"
            cel.Offset(0, -38).Value = "Intermediate Stage Architectural Drawings"

       ElseIf InStr(1, cel.Value, "TECHNICAL INTERFACE") > 0 Then
            cel.Offset(0, -39).Value = "Architectural"
            cel.Offset(0, -38).Value = "Technical Interfaces"

       ElseIf InStr(1, cel.Value, "COMMISSIONING") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Commissioning Documents"

       ElseIf InStr(1, cel.Value, "COMMISSIONING REPORT") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Commissioning Records & Reports"
			
       ElseIf InStr(1, cel.Value, "COMMISSIONING RECORD") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Commissioning Records & Reports"

       ElseIf InStr(1, cel.Value, "CONSTRUCTION PERMITS/LICENCE/PERMIT PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Construction Permits/Licence/Permit Plan"

       ElseIf InStr(1, cel.Value, "MINUTES OF ALL CONTRACTORS MEETING") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Minutes of All Contractor's Meetings"

       ElseIf InStr(1, cel.Value, "WORK CHANGE PROPOSAL") > 0 Then
            cel.Offset(0, -39).Value = "Construction & Commissioning"
            cel.Offset(0, -38).Value = "Work Change Proposal"

       ElseIf InStr(1, cel.Value, "AGGREGATE TESTING REPORT") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Aggregate Testing Reports"

       ElseIf InStr(1, cel.Value, "ALL STRUCTURAL & CIVIL DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "All Structural & Civil Drawings"
            
       ElseIf InStr(1, cel.Value, "STRUCTURAL") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "All Structural & Civil Drawings"

       ElseIf InStr(1, cel.Value, "CEMENT MTOS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Cement MTOs"

       ElseIf InStr(1, cel.Value, "CERTIFIED MILL TEST REPORTS FOR STRUCTURAL STEEL") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Certified Mill Test Reports for Structural Steel"

       ElseIf InStr(1, cel.Value, "CIVIL MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Civil Miscellaneous"

       ElseIf InStr(1, cel.Value, "CIVIL SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Civil Specifications"

       ElseIf InStr(1, cel.Value, "COMPACTION TESTS FOR EARTH WORK") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Compaction Tests for Earth Work"

       ElseIf InStr(1, cel.Value, "CONCRETE AND STRUCTURAL SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Concrete and Structural Specifications"

       ElseIf InStr(1, cel.Value, "CONCRETE APPLICATION PROCEDURE") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Concrete Application Procedure"

       ElseIf InStr(1, cel.Value, "CONCRETE CURING PROCEDURE") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Concrete Curing Procedure"

       ElseIf InStr(1, cel.Value, "CONCRETE MIX DESIGN") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Concrete Mix Design"

       ElseIf InStr(1, cel.Value, "CONCRETE TEST RESULTS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Concrete Test Results"

       ElseIf InStr(1, cel.Value, "CULVERTS, BUND WALLS & SUMPS DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Culverts, Bund Walls & Sumps Drawings"

       ElseIf InStr(1, cel.Value, "DESIGN INTEGRITY DOCUMENTS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Design Integrity Documents"

       ElseIf InStr(1, cel.Value, "DRAINAGE & STORM WATER SEWER SYSTEM DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Drainage & Storm Water Sewer System Drawings"

       ElseIf InStr(1, cel.Value, "FABRICATION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Fabrication Drawings"

       ElseIf InStr(1, cel.Value, "FIREPROOFING DRAWINGS & SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Fireproofing Drawings & Specification"

       ElseIf InStr(1, cel.Value, "FOUNDATION DRAWINGS/DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Foundation Drawings/Details"
			
       ElseIf InStr(1, cel.Value, "FOUNDATION DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Foundation Drawings/Details"
			
       ElseIf InStr(1, cel.Value, "FOUNDATION DETAIL") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Foundation Drawings/Details"
			
       ElseIf InStr(1, cel.Value, "FOUNDATION") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Foundation Drawings/Details"
			
       ElseIf InStr(1, cel.Value, "GEOTECHNICAL & GEOPHYSICAL INVESTIGATION SERVICE WORK PLAN, REPORTS & MAPS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Geotechnical & Geophysical Investigation Service Work Plan, Reports & Maps"

       ElseIf InStr(1, cel.Value, "GRADING PLAN, LEVELING") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Grading Plan, Leveling"

       ElseIf InStr(1, cel.Value, "IRRIGATION/LANDSCAPING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Irrigation/Landscaping Drawings"

       ElseIf InStr(1, cel.Value, "MANUFACTURERS CERTIFICATION FOR BOLTS, NUTS AND WASHERS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Manufacturer's Certification for Bolts, Nuts and Washers"

       ElseIf InStr(1, cel.Value, "PAVEMENT, SURFACING & SLOPE PROTECTION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Pavement, Surfacing & Slope Protection Drawings"

       ElseIf InStr(1, cel.Value, "ROAD, FENCING & DRAINAGE LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Road, Fencing & Drainage Layout"

       ElseIf InStr(1, cel.Value, "SAND TESTING REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Sand Testing Reports"

       ElseIf InStr(1, cel.Value, "SCHEDULES & BILL OF MATERIALS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Schedules & Bill Of Materials"

       ElseIf InStr(1, cel.Value, "SHOP DETAIL DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Shop Detail Drawings"

       ElseIf InStr(1, cel.Value, "SITE SURVEY") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Site Survey (Data & Information Report)"

       ElseIf InStr(1, cel.Value, "STRUCTURAL STEEL CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Structural Steel Calculations"

       ElseIf InStr(1, cel.Value, "TOPOGRAPHIC DRAWING OF SITE") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Topographic Drawing of Site"

       ElseIf InStr(1, cel.Value, "TRENCHING AND UNDERGROUND SERVICES LAYOUTS AND CONCRETE DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Civil"
            cel.Offset(0, -38).Value = "Trenching and Underground Services Layouts and Concrete Details"

       ElseIf InStr(1, cel.Value, "BLOCK DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Block Diagrams"

       ElseIf InStr(1, cel.Value, "BUILDING ELECTRICAL LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Building Electrical Layouts"

       ElseIf InStr(1, cel.Value, "BUILDING MANAGEMENT SYSTEM DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Building Management System Drawings (BMS)"

       ElseIf InStr(1, cel.Value, "BUILDING SMALL POWER AND LIGHTING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Building Small Power and Lighting Layouts"

       ElseIf InStr(1, cel.Value, "CABLE SIZE CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Cable Size Calculations"

       ElseIf InStr(1, cel.Value, "CATHODIC PROTECTION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Cathodic Protection Drawings"

       ElseIf InStr(1, cel.Value, "CERTIFICATES, REPORTS & TEST PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Certificates, Reports & Test Plan"

       ElseIf InStr(1, cel.Value, "CONDUIT SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Conduit Schedule"

       ElseIf InStr(1, cel.Value, "EARTHING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Earthing Layouts"

       ElseIf InStr(1, cel.Value, "ELECTRICAL CABLE BLOCK DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Cable Block Diagrams"

       ElseIf InStr(1, cel.Value, "ELECTRICAL CABLE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Cable Schedule"

       ElseIf InStr(1, cel.Value, "ELECTRICAL DRAWINGS FOR BUILDING SERVICES") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Drawings For Building Services"

       ElseIf InStr(1, cel.Value, "ELECTRICAL EQUIP. LOCATION & CABLE ROUTING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Equip. Location & Cable Routing Layouts"

       ElseIf InStr(1, cel.Value, "ELECTRICAL EQUIPMENT ARRANGEMENT LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Equipment Arrangement Layouts"

       ElseIf InStr(1, cel.Value, "ELECTRICAL INSTALLATION STANDARDS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Installation Standards (Power, Lighting, Cable Trench Sections, Earthing etc.)"

       ElseIf InStr(1, cel.Value, "ELECTRICAL INTERCONNECTION DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Interconnection Diagrams"

       ElseIf InStr(1, cel.Value, "ELECTRICAL LOAD ANALYSIS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Load Analysis"

       ElseIf InStr(1, cel.Value, "ELECTRICAL LOGIC DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Logic Diagrams"

       ElseIf InStr(1, cel.Value, "ELECTRICAL MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Miscellaneous"

       ElseIf InStr(1, cel.Value, "ELECTRICAL PANEL SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Panel Schedule"

       ElseIf InStr(1, cel.Value, "ELECTRICAL SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical Specifications"

       ElseIf InStr(1, cel.Value, "ELECTRICAL SYSTEM STUDIES") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Electrical System Studies"

       ElseIf InStr(1, cel.Value, "GROUNDING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Grounding Drawings"

       ElseIf InStr(1, cel.Value, "LEGENDS AND SYMBOL LISTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Legends and Symbol Lists"

       ElseIf InStr(1, cel.Value, "LIGHTING AND GROUNDING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Lighting and Grounding Layouts"

       ElseIf InStr(1, cel.Value, "LIGHTING FIXTURE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Lighting Fixture Schedule"

       ElseIf InStr(1, cel.Value, "BILL OF MATERIAL") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Lists/Bill Of Materials"

       ElseIf InStr(1, cel.Value, "OVERHEAD TRANSMISSION LINES DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Overhead Transmission Lines Drawings"

       ElseIf InStr(1, cel.Value, "POWER AND GROUNDING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Power and Grounding Drawings"
			
       ElseIf InStr(1, cel.Value, "GROUNDING") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Power and Grounding Drawings"

       ElseIf InStr(1, cel.Value, "POWER SYSTEM AND LOAD TRANSFER PHILOSOPHY") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Power System and Load Transfer Philosophy"

       ElseIf InStr(1, cel.Value, "PROTECTION RELAY COORDINATION SETTINGS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Protection Relay Coordination Settings (Calculations, Curves & Tables)"

       ElseIf InStr(1, cel.Value, "PROTECTIVE DEVICE CO-ORDINATION") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Protective Device Co-ordination"

       ElseIf InStr(1, cel.Value, "SCHEMATIC WIRING DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Schematic Wiring Diagrams"

       ElseIf InStr(1, cel.Value, "SHORT CIRCUIT STUDIES") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Short Circuit Studies"

       ElseIf InStr(1, cel.Value, "SINGLE LINE DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Single Line Diagrams"
			
       ElseIf InStr(1, cel.Value, "SLD") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Single Line Diagrams"

       ElseIf InStr(1, cel.Value, "STANDARD CABLE TRAY DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Standard Cable Tray Details"

       ElseIf InStr(1, cel.Value, "STANDARD GROUND DETAIL") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Standard Ground Details"

       ElseIf InStr(1, cel.Value, "STANDARD LIGHTING DETAIL") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Standard Lighting Details"

       ElseIf InStr(1, cel.Value, "STANDARD POWER DETAIL") > 0 Then
            cel.Offset(0, -39).Value = "Electrical"
            cel.Offset(0, -38).Value = "Standard Power Details"

       ElseIf InStr(1, cel.Value, "DETAILED PROCEDURES FOR F&G SYSTEM FAT") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "Detailed Procedures for F&G System FAT"

       ElseIf InStr(1, cel.Value, "FAILURE MODES EFFECTS ANALYSIS FOR F&G SYSTEM") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "Failure Modes Effects Analysis for F&G Systems"

       ElseIf InStr(1, cel.Value, "F&G CABLE BLOCK DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G Cable Block Diagrams"

       ElseIf InStr(1, cel.Value, "F&G CABLE ROUTING LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G Cable Routing Layouts"

       ElseIf InStr(1, cel.Value, "F&G DETECTORS LOCATION") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G Detectors Location/Cable Routing Layouts (Including Building Services)"

       ElseIf InStr(1, cel.Value, "F&G JUNCTION BOX SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G Junction Box Schedule"

       ElseIf InStr(1, cel.Value, "F&G SCHEMATIC") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G Schematic"

       ElseIf InStr(1, cel.Value, "F&G SYSTEM POWER DISTRIBUTION, GROUNDING INTERCONNECTS DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "F&G System Power Distribution, Grounding Interconnects Drawing"

       ElseIf InStr(1, cel.Value, "OVERALL F&G SYSTEM BLOCK DIAGRAMS, I/O LOADING & CABINET LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "Overall F&G System Block Diagrams, I/O Loading & Cabinet Layouts"

       ElseIf InStr(1, cel.Value, "STANDARD F&G DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Fire & Gas"
            cel.Offset(0, -38).Value = "Standard F&G Details"

       ElseIf InStr(1, cel.Value, "ADVANCE REVISION NOTICE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Advance Revision Notice (ARN)"

       ElseIf InStr(1, cel.Value, "CHANGE CONTROL RECORD") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Change Control Records"

       ElseIf InStr(1, cel.Value, "CONTRACTORS INTENTION OF MAJOR CHANGE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Contractor's Intention of Major Change"

       ElseIf InStr(1, cel.Value, "DESIGN CHANGE CONTROL PROCEDURE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Design Change Control Procedures"

       ElseIf InStr(1, cel.Value, "DESIGN DEVIATION REPORT") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Design Deviation Reports"

       ElseIf InStr(1, cel.Value, "DESIGN GENERAL SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Design General Specifications"

       ElseIf InStr(1, cel.Value, "DOCUMENT DISTRIBUTION SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Document Distribution Schedule"

       ElseIf InStr(1, cel.Value, "DRAFTING STANDARD") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Drafting Standards"

       ElseIf InStr(1, cel.Value, "MINUTES OF MEETING") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Minutes of Meetings (MOM)/Record of Conversation (ROC)"

       ElseIf InStr(1, cel.Value, "PROGRESS REPORT") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Monthly/Other Progress Reports"

       ElseIf InStr(1, cel.Value, "PROJECT ENGINEERING PROCEDURE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Project Engineering Procedure"

       ElseIf InStr(1, cel.Value, "PROJECT SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Project Specifications (New, not covered by DGS)"

       ElseIf InStr(1, cel.Value, "REPORTS & RECORD") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Reports & Records"

       ElseIf InStr(1, cel.Value, "REVIEW AND APPROVAL") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Review and Approval"

       ElseIf InStr(1, cel.Value, "TECHNICAL QUERY FORM") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Technical Query (TQ) Form"

       ElseIf InStr(1, cel.Value, "WORK BREAKDOWN STRUCTURE") > 0 Then
            cel.Offset(0, -39).Value = "General"
            cel.Offset(0, -38).Value = "Work Breakdown Structure (WBS)"

       ElseIf InStr(1, cel.Value, "ACCIDENT REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Accident Reports"

       ElseIf InStr(1, cel.Value, "ACTIVE FIRE FIGHTING SYSTEM") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Active Fire Fighting System P&ID"

       ElseIf InStr(1, cel.Value, "ACTIVE FIRE PROTECTION SYSTEM DATA") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Active Fire Protection System Data"

       ElseIf InStr(1, cel.Value, "COMMUNITY AFFAIRS PLAN") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Community Affairs Plan"

       ElseIf InStr(1, cel.Value, "EMERGENCY PROCEDURES MANUALS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Emergency Procedures Manuals"

       ElseIf InStr(1, cel.Value, "ENVIRONMENTAL MANAGEMENT PLAN") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Environmental Management Plan"

       ElseIf InStr(1, cel.Value, "EQUIPMENT FAILURE INFORMATION") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Equipment Failure Information"

       ElseIf InStr(1, cel.Value, "FIRE FIGHTING EQUIPMENT LAYOUT DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Fire Fighting Equipment Layout Drawings"

       ElseIf InStr(1, cel.Value, "FIRE PROOFING SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Fire Proofing Schedules"

       ElseIf InStr(1, cel.Value, "FIRE WATER SYSTEM LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Fire Water System Layout"

       ElseIf InStr(1, cel.Value, "FIRE WATER SYSTEM SCHEMATIC") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Fire Water System Schematic"

       ElseIf InStr(1, cel.Value, "HAZARDOUS AREA CLASSIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Hazardous Area Classification"

       ElseIf InStr(1, cel.Value, "HAZARDOUS SOURCE DATA") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Hazardous Source Data"

       ElseIf InStr(1, cel.Value, "HAZARDOUS SUBSTANCE DATA SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Hazardous Substance Data Sheets"

       ElseIf InStr(1, cel.Value, "HAZOP REPORT AND PROJECT RESPONSE") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "HAZOP Report and Project Response"

       ElseIf InStr(1, cel.Value, "HAZOP REVIEW AND REPORT") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "HAZOP Review and Report"

       ElseIf InStr(1, cel.Value, "HEALTH/SAFETY/ENVIRONMENT PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Health/Safety/Environment Procedures"

       ElseIf InStr(1, cel.Value, "HIPS DOSSIER") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "HIPS Dossier"

       ElseIf InStr(1, cel.Value, "LIFE SAVING EQUIPMENT & ESCAPE ROUTE DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Life Saving Equipment & Escape Route Drawing"

       ElseIf InStr(1, cel.Value, "LIST OF PROPOSED DISPOSAL SITES") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "List of Proposed Disposal Sites"

       ElseIf InStr(1, cel.Value, "MONTHLY REPORT - HEALTH & SAFETY") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Monthly Report - Health & Safety"

       ElseIf InStr(1, cel.Value, "PASSIVE & ACTIVE FIRE PROTECTION PHILOSOPHY") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Passive & Active Fire Protection Philosophy"

       ElseIf InStr(1, cel.Value, "PASSIVE FIRE PROTECTION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Passive Fire Protection Drawings"

       ElseIf InStr(1, cel.Value, "PHSR REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "PHSR Reports"

       ElseIf InStr(1, cel.Value, "RECOMMENDATION MONITORING SYSTEM") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Recommendation Monitoring System"

       ElseIf InStr(1, cel.Value, "RECORD OF SURVEY") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Record of Survey"

       ElseIf InStr(1, cel.Value, "REGISTER OF SAFETY DEVICES") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Register of Safety Devices"

       ElseIf InStr(1, cel.Value, "SAFETY EQUIPMENT DATA SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety Equipment Data Sheets"

       ElseIf InStr(1, cel.Value, "SAFETY EQUIPMENT DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety Equipment Drawings"

       ElseIf InStr(1, cel.Value, "SAFETY EQUIPMENT LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety Equipment Layout"

       ElseIf InStr(1, cel.Value, "SAFETY/FIRE FIGHTING MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety/Fire Fighting Miscellaneous"

       ElseIf InStr(1, cel.Value, "SAFETY/HEALTH PROGRAMME") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety/Health Programme"

       ElseIf InStr(1, cel.Value, "SAFETY SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Safety Specifications"

       ElseIf InStr(1, cel.Value, "SIL REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "SIL Reports"

       ElseIf InStr(1, cel.Value, "TEST PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Test Procedures"

       ElseIf InStr(1, cel.Value, "WASTE MATERIAL MANAGEMENT PLAN") > 0 Then
            cel.Offset(0, -39).Value = "HSE"
            cel.Offset(0, -38).Value = "Waste Material Management Plan"

       ElseIf InStr(1, cel.Value, "ALL CERTIFIED EQUIPMENT AND ACCESSORIES DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "All Certified Equipment and Accessories Drawings"

       ElseIf InStr(1, cel.Value, "ALL HVAC-RELATED CONSTRUCTION SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "All HVAC-related Construction Specifications"

       ElseIf InStr(1, cel.Value, "CONTROL LOGIC DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Control Logic Diagrams"

       ElseIf InStr(1, cel.Value, "CONTROL SCHEMATIC DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Control Schematic Diagrams"

       ElseIf InStr(1, cel.Value, "CONTROL TERMINATION DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Control Termination Diagrams"

       ElseIf InStr(1, cel.Value, "COOLING, HEATING, VENTILATION LOAD CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Cooling, Heating, Ventilation Load Calculations"

       ElseIf InStr(1, cel.Value, " DUCTWORK") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Drawings (Equipment, Ductwork, Piping & Accessories)"

       ElseIf InStr(1, cel.Value, "DUCTWORK - MATERIAL SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Ductwork - Material Specification"

       ElseIf InStr(1, cel.Value, "EQUIPMENT SCHEDULE WITH CAPACITIES") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Equipment Schedule with Capacities"

       ElseIf InStr(1, cel.Value, "FILTERS AND RETAINER FRAMES DATA") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Filters and Retainer Frames Data"

       ElseIf InStr(1, cel.Value, "DIFFUSER PERFORMANCE DATA") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Grilles, Registers and Diffuser Performance Data"

       ElseIf InStr(1, cel.Value, "HVAC DUCT") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "HVAC Duct and Instrument Diagrams"

       ElseIf InStr(1, cel.Value, "HVAC GENERAL ARRANGEMENT DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "HVAC General Arrangement Drawings"

       ElseIf InStr(1, cel.Value, "INSULATION DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Insulation Details"

       ElseIf InStr(1, cel.Value, "PERFORMANCE DATA AND OPERATIONAL CURVES") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Performance Data and Operational Curves"

       ElseIf InStr(1, cel.Value, "PIPING SYSTEM VALVES, FITTINGS & ACCESSORIES DRAWING DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "HVAC"
            cel.Offset(0, -38).Value = "Piping System Valves, Fittings & Accessories Drawing Details"

       ElseIf InStr(1, cel.Value, "AUXILIARY RACKS LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Auxiliary Racks Layouts"

       ElseIf InStr(1, cel.Value, "CABLE BLOCK DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Cable Block Diagrams"

       ElseIf InStr(1, cel.Value, "CABLE ROUTING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Cable Routing Drawings"

       ElseIf InStr(1, cel.Value, "CABLE ROUTING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Cable Routing Layouts"

       ElseIf InStr(1, cel.Value, "CAUSE AND EFFECT DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Cause and Effect Diagrams"

       ElseIf InStr(1, cel.Value, "CONTROL PANEL LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Control Panel Layouts"

       ElseIf InStr(1, cel.Value, "CONTROL ROOM LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Control Room Layouts"

       ElseIf InStr(1, cel.Value, "CONTROL SCHEMATIC") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Control Schematic"

       ElseIf InStr(1, cel.Value, "CRITICAL INSTRUMENT AND DESIGN SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Critical Instrument and Design Specification"

       ElseIf InStr(1, cel.Value, "CRITICAL INSTRUMENT CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Critical Instrument Calculations"

       ElseIf InStr(1, cel.Value, "DETAILED PROCEDURES FOR CONTROL AND SHUTDOWN SYSTEM FAT") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Detailed Procedures for Control and Shutdown System FAT"

       ElseIf InStr(1, cel.Value, "ESD DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "ESD Diagrams"

       ElseIf InStr(1, cel.Value, "FAILURE MODES EFFECTS ANALYSIS FOR SHUTDOWN SYSTEMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Failure Modes Effects Analysis for Shutdown Systems"

       ElseIf InStr(1, cel.Value, "FISCAL METERING DESIGN CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Fiscal Metering Design Calculations"

       ElseIf InStr(1, cel.Value, "FUNCTIONAL SPECIFICATION FOR SYSTEM CONTROL APPLICATIONS INCLUDING STANDARD PROCESS CONTROL AND APC") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Functional Specification for System Control Applications Including Standard Process Control and APC"

       ElseIf InStr(1, cel.Value, "HOOK-UP DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Hook-Up Drawings"

       ElseIf InStr(1, cel.Value, "INSTALLATION AND MOUNTING STANDARDS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Installation and Mounting Standards"

       ElseIf InStr(1, cel.Value, "INSTRUMENT AND DESIGN SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument and Design Specifications"

       ElseIf InStr(1, cel.Value, "INSTRUMENT CABLE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Cable Schedule"

       ElseIf InStr(1, cel.Value, "INSTRUMENT/CONTROL SYSTEM POWER DISTRIBUTION, GROUNDING INTERCONNECTS DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument/Control System Power Distribution, Grounding Interconnects Drawing"

       ElseIf InStr(1, cel.Value, "INSTRUMENT DATASHEET") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Datasheet"

       ElseIf InStr(1, cel.Value, "INSTRUMENT GENERAL ARRANGEMENT DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument General Arrangement Drawings"

       ElseIf InStr(1, cel.Value, "INSTRUMENT INDEX, I/O LIST & ALARM LIST") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Index, I/O List & Alarm List"

       ElseIf InStr(1, cel.Value, "INSTRUMENT INSTALLATION STANDARDS : CABLE TRENCH") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Installation Standards : Cable Trench (Sections/Details)"

       ElseIf InStr(1, cel.Value, "INSTRUMENT JUNCTION BOX SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Junction Box Schedule"

       ElseIf InStr(1, cel.Value, "INSTRUMENT LOCATION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Location Drawings\Layouts"

       ElseIf InStr(1, cel.Value, "INSTRUMENT LOOP DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Loop Drawings"

       ElseIf InStr(1, cel.Value, "INSTRUMENT WIRING & TERMINATION DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Instrument Wiring & Termination Diagrams"

       ElseIf InStr(1, cel.Value, "INTERFACE DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Interface Drawings"

       ElseIf InStr(1, cel.Value, "I/O LOADING LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "I/O Loading Layouts"

       ElseIf InStr(1, cel.Value, "BILL OF MATERIALS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Lists/Bill Of Materials"

       ElseIf InStr(1, cel.Value, "LOGIC DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Logic Diagrams"

       ElseIf InStr(1, cel.Value, "MARSHALLING CABINETS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Marshalling Cabinets"

       ElseIf InStr(1, cel.Value, "OVERALL CONTROL SYSTEM BLOCK DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Overall Control System Block Diagrams"

       ElseIf InStr(1, cel.Value, "RELAY SCHEMATICS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Relay Schematics"

       ElseIf InStr(1, cel.Value, "ROOM LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Room Layout"

       ElseIf InStr(1, cel.Value, "RELIEF VALVE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Safety/Relief Valve Schedule"

       ElseIf InStr(1, cel.Value, "SCADA DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "SCADA/PLC/DCS Details"

       ElseIf InStr(1, cel.Value, "SCHEMATIC DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Schematic Diagram"

       ElseIf InStr(1, cel.Value, "SHUTDOWN LOGIC DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Shutdown Logic Diagrams"

       ElseIf InStr(1, cel.Value, "SOFTWARE CONFIGURATION DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Software Configuration Diagrams"

       ElseIf InStr(1, cel.Value, "STANDARD INSTRUMENT DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Standard Instrument Details"

       ElseIf InStr(1, cel.Value, "TESTING AND COMMISSIONING PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Instrumentation"
            cel.Offset(0, -38).Value = "Testing and Commissioning Procedures"

       ElseIf InStr(1, cel.Value, "ACCESS MANUALS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Access Manuals"

       ElseIf InStr(1, cel.Value, "ALIGNMENT SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Alignment Specifications"

       ElseIf InStr(1, cel.Value, "CONDITION MONITORING BASELINES") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Condition Monitoring Baselines"

       ElseIf InStr(1, cel.Value, "CRITICAL WELD PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Critical Weld Procedures"

       ElseIf InStr(1, cel.Value, "EQUIPMENT LAYOUTS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Equipment Layouts"

       ElseIf InStr(1, cel.Value, "EQUIPMENT MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Equipment Miscellaneous"

       ElseIf InStr(1, cel.Value, "LIFTING EQUIPMENT CERTIFICATE") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Lifting Equipment Certificate"

       ElseIf InStr(1, cel.Value, "LUBRICATION SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Lubrication Schedule"

       ElseIf InStr(1, cel.Value, "MECHANICAL EQUIPMENT DATA SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Mechanical Equipment Data Sheets"

       ElseIf InStr(1, cel.Value, "EQUIPMENT  GENERAL ARRANGEMENT DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Mechanical\Equipment  General Arrangement Drawing"

       ElseIf InStr(1, cel.Value, "PERFORMANCE CURVES") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Performance Curves"

       ElseIf InStr(1, cel.Value, "PLOT PLANS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Plot Plans"

       ElseIf InStr(1, cel.Value, "PRESSURE VESSEL CERTIFICATE") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Pressure Vessel Certificate"

       ElseIf InStr(1, cel.Value, "PSV TEST CERTIFICATE") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "PSV Test Certificate"

       ElseIf InStr(1, cel.Value, "STANDARD WELD PROCEDURE") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Standard Weld Procedure"

       ElseIf InStr(1, cel.Value, "START-UP LOGS FOR MAJOR ROTATING MACHINERY") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Start-up Logs for major rotating machinery"

       ElseIf InStr(1, cel.Value, "VESSEL DESIGN CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Vessel Design Calculations"

       ElseIf InStr(1, cel.Value, "VESSELS & TANKS DATA SHEET") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Vessels & Tank’s Data Sheet"

       ElseIf InStr(1, cel.Value, "VESSELS & TANKS SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Vessels & Tank’s Specifications"

       ElseIf InStr(1, cel.Value, "VESSELS & TANKS DETAIL DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Vessels & Tank's Detail Drawing"

       ElseIf InStr(1, cel.Value, "VESSELS & TANKS MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "Mechanical"
            cel.Offset(0, -38).Value = "Vessels & Tanks Miscellaneous"

       ElseIf InStr(1, cel.Value, "BASELINE CORROSION READINGS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Baseline Corrosion Readings"

       ElseIf InStr(1, cel.Value, "CATHODIC PROTECTION FIELD TESTING REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Cathodic Protection Field Testing Reports"

       ElseIf InStr(1, cel.Value, "COATINGS TEST REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Coatings Test Reports"

       ElseIf InStr(1, cel.Value, "CORROSION CONTROL PHILOSOPHY") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Corrosion Control Philosophy"

       ElseIf InStr(1, cel.Value, "CORROSION MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Corrosion Miscellaneous"

       ElseIf InStr(1, cel.Value, "MATERIAL SELECTION DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Material Selection Drawings"

       ElseIf InStr(1, cel.Value, "MATERIAL SELECTION PHILOSOPHY") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Material Selection Philosophy"

       ElseIf InStr(1, cel.Value, "PAINTING SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Material Management"
            cel.Offset(0, -38).Value = "Painting Specifications"

       ElseIf InStr(1, cel.Value, "ENGINEERING REPORTS & CALCULATION PACKAGES") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Engineering Reports & Calculation Packages"

       ElseIf InStr(1, cel.Value, "MAINTENANCE AND PRESERVATION PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Maintenance and Preservation Procedures"

       ElseIf InStr(1, cel.Value, "MAINTENANCE INFORMATION") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Maintenance Information"

       ElseIf InStr(1, cel.Value, "OPERATING INFORMATION") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Operating Information"

       ElseIf InStr(1, cel.Value, "SPARES LISTINGS") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Spares Listings"

       ElseIf InStr(1, cel.Value, "SUPPLIER DATA DOSSIERS") > 0 Then
            cel.Offset(0, -39).Value = "Operate and Maintain"
            cel.Offset(0, -38).Value = "Supplier Data Dossiers"

       ElseIf InStr(1, cel.Value, "CONTRACTORS FORM FORMATS") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Contractor's Form Formats"

       ElseIf InStr(1, cel.Value, "CONTRACTORS LIST OF SUBCONTRACTED WORK") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Contractor's List of Subcontracted Work"

       ElseIf InStr(1, cel.Value, "CONTRACTORS PROPOSED SUPPLIER LIST") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Contractor's Proposed Supplier List"

       ElseIf InStr(1, cel.Value, "MATERIAL FINAL RECONCILIATION LIST") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Final Reconciliation List"

       ElseIf InStr(1, cel.Value, "MATERIAL MANAGEMENT PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Management Procedures"

       ElseIf InStr(1, cel.Value, "MATERIAL MOVEMENT REQUEST") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Movement Request (MMR)"

       ElseIf InStr(1, cel.Value, "MATERIAL MOVEMENT TICKET") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Movement Ticket (MMT)"

       ElseIf InStr(1, cel.Value, "MATERIAL RECEIVING REPORT") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Receiving Report (MRR)"

       ElseIf InStr(1, cel.Value, "MATERIAL RELEASES") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Releases"

       ElseIf InStr(1, cel.Value, "MATERIAL STATUS REPORTING") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Status Reporting"

       ElseIf InStr(1, cel.Value, "OFF-SITE TRANSFER AUTHORISATION") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Off-site Transfer Authorisation (OST)"

       ElseIf InStr(1, cel.Value, "OVER SHORT & DAMAGED REPORT") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Over Short & Damaged Report (OSD)"

       ElseIf InStr(1, cel.Value, "PROCUREMENT INSPECTION - SURVEILLANCE PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Procurement Inspection - Surveillance Plan"

       ElseIf InStr(1, cel.Value, "PURCHASING PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Purchasing Plan / Procurement Procedures"

       ElseIf InStr(1, cel.Value, "SCRAP DISPOSAL REQUISITION") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Scrap Disposal Requisition"

       ElseIf InStr(1, cel.Value, "SURPLUS MATERIAL REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Surplus Material Reports"

       ElseIf InStr(1, cel.Value, "UNPRICED COPY OF PURCHASE ORDERS") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Unpriced Copy of Purchase Orders (as issued)"

       ElseIf InStr(1, cel.Value, "UNPRICED COPY OF SUBCONTRACTS") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Unpriced Copy of Subcontracts (as issued)"

       ElseIf InStr(1, cel.Value, "CROSSING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Crossing Drawings"

       ElseIf InStr(1, cel.Value, "PIPELINE ALIGNMENT SHEETS,") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline Alignment Sheets,"

       ElseIf InStr(1, cel.Value, "PIPELINE CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline Calculations"

       ElseIf InStr(1, cel.Value, "PIPELINE DATA SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline Data Sheets"

       ElseIf InStr(1, cel.Value, "PIPELINE GENERAL ARRANGEMENT DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Pipeline General Arrangement Drawings"

       ElseIf InStr(1, cel.Value, "ROUTE MAP") > 0 Then
            cel.Offset(0, -39).Value = "Pipeline"
            cel.Offset(0, -38).Value = "Route Maps"

       ElseIf InStr(1, cel.Value, "CATHODIC PROTECTION DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Cathodic Protection Drawing"

       ElseIf InStr(1, cel.Value, "COMPLIANCE DOCUMENTS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Compliance Documents"

       ElseIf InStr(1, cel.Value, "DESIGN CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Design Calculations"

       ElseIf InStr(1, cel.Value, "DESIGN, FABRICATION AND INSTALLATION SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Design, Fabrication and Installation Specifications"

       ElseIf InStr(1, cel.Value, "EMERGENCY P/L VALVE CERTIFICATE") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Emergency P/L Valve Certificate"

       ElseIf InStr(1, cel.Value, "FABRICATION ISOMETRIC DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Fabrication Isometric Drawing"

       ElseIf InStr(1, cel.Value, "FABRICATION RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Fabrication Records"

       ElseIf InStr(1, cel.Value, "INSTALLATION PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Installation Procedures"

       ElseIf InStr(1, cel.Value, "INSULATION SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Insulation Specification"

       ElseIf InStr(1, cel.Value, "LINE LIST & LINE SCHEDULE DATA") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Line List & Line Schedule Data"

       ElseIf InStr(1, cel.Value, "PIPE STRESS PACK") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Pipe Stress Pack"

       ElseIf InStr(1, cel.Value, "PIPING LAYOUTS & PLOT PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Piping Layouts & Plot Plan"

       ElseIf InStr(1, cel.Value, "PIPING SPECIAL ITEMS SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Piping Special Items Specification"

       ElseIf InStr(1, cel.Value, "PIPING SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Piping Specification"

       ElseIf InStr(1, cel.Value, "PIPING SUPPORT STANDARDS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Piping Support Standards"

       ElseIf InStr(1, cel.Value, "PRESSURE TEST RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Pressure Test Records"

       ElseIf InStr(1, cel.Value, "SCHEDULES & BILL OF MATERIALS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Schedules & Bill Of Materials"

       ElseIf InStr(1, cel.Value, "STRESS ANALYSIS REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Stress Analysis Reports"

       ElseIf InStr(1, cel.Value, "TIE-IN/BATTERY LIMIT SCHEDULES") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Tie-in/Battery Limit Schedules"

       ElseIf InStr(1, cel.Value, "VALVE SPECIFICATION") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Valve Specification"

       ElseIf InStr(1, cel.Value, "WATER ANALYSIS - CERT TEST REPORT") > 0 Then
            cel.Offset(0, -39).Value = "Piping"
            cel.Offset(0, -38).Value = "Water Analysis - Cert Test Report"

       ElseIf InStr(1, cel.Value, "ALL PROJECT PHILOSOPHIES") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "All Project Philosophies"

       ElseIf InStr(1, cel.Value, "ALL PROJECT STUDIES") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "All Project Studies (generated in the EPC phase and later)"

       ElseIf InStr(1, cel.Value, "CAUSE AND EFFECT CHARTS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Cause and Effect Charts"

       ElseIf InStr(1, cel.Value, "CRITICALITY RESUME") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Criticality Resume"

       ElseIf InStr(1, cel.Value, "EQUIPMENT LIST ") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Equipment List "

       ElseIf InStr(1, cel.Value, "FACILITY BLOCK DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Facility Block Diagram"

       ElseIf InStr(1, cel.Value, "FLUID AND SUBSTANCE DATA") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Fluid and Substance Data"

       ElseIf InStr(1, cel.Value, "HL/LP INTERFACE DETAILS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "HL/LP Interface Details"

       ElseIf InStr(1, cel.Value, "MASS BALANCE") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Mass Balance"

       ElseIf InStr(1, cel.Value, "MATERIAL SELECTION DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Material Selection Diagram"

       ElseIf InStr(1, cel.Value, "POWER AND UTILITY SUMMARY") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Power and Utility Summary"

       ElseIf InStr(1, cel.Value, "PROCESS CALCULATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Calculations"

       ElseIf InStr(1, cel.Value, "PFD") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Flow Diagrams (PFD)"

       ElseIf InStr(1, cel.Value, "PROCESS SAFEGUARDING DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Safeguarding Drawings"

       ElseIf InStr(1, cel.Value, "PROCESS SIMULATION MODEL") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Process Simulation Model"

       ElseIf InStr(1, cel.Value, "RELIEF AND BLOW DOWN DATA") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Relief and Blow Down Data"

       ElseIf InStr(1, cel.Value, "SUPERVISORY LEVEL OPERATING PROCEDURES FOR ALL PROCESS/UTILITY SYSTEMS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Supervisory Level Operating Procedures for all Process/Utility Systems"

       ElseIf InStr(1, cel.Value, "SURGE AND DYNAMIC ANALYSIS DATA") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Surge and Dynamic Analysis Data"

       ElseIf InStr(1, cel.Value, "SYSTEM DESIGN BASIS REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "System Design Basis Reports"

       ElseIf InStr(1, cel.Value, "UFD") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Utility Flow Diagrams (UFD)"

       ElseIf InStr(1, cel.Value, "UTILITY SUMMARY") > 0 Then
            cel.Offset(0, -39).Value = "Process"
            cel.Offset(0, -38).Value = "Utility Summary"

       ElseIf InStr(1, cel.Value, "CONSTRUCTION, INSPECTION & TEST PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Construction, Inspection & Test Plan"

       ElseIf InStr(1, cel.Value, "EQUIPMENT CRITICALITY RATING SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Equipment Criticality Rating Sheets"

       ElseIf InStr(1, cel.Value, "INSPECTION MANNING LEVELS BY JOB FUNCTION/LOCATION") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Inspection Manning Levels by Job Function/Location"

       ElseIf InStr(1, cel.Value, "MATERIAL TRACE ABILITY REPORT/RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Material Trace ability Report/Records"

       ElseIf InStr(1, cel.Value, "NDE PERSONNEL QUALIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "NDE Personnel Qualifications"

       ElseIf InStr(1, cel.Value, "NDE PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "NDE Procedures"

       ElseIf InStr(1, cel.Value, "NDE TEST REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "NDE Test Reports"

       ElseIf InStr(1, cel.Value, "PRESSURE TESTING PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Pressure Testing Procedures"

       ElseIf InStr(1, cel.Value, "QA/QC GROUP ORGANISATION CHARTS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "QA/QC Group Organisation Charts"

       ElseIf InStr(1, cel.Value, "QUALITY AUDIT REPORTS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Quality Audit Reports"

       ElseIf InStr(1, cel.Value, "QUALITY AUDIT SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Quality Audit Schedule"

       ElseIf InStr(1, cel.Value, "QUALITY MANUAL/QUALITY SYSTEM PROGRAMME") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Quality Manual/Quality System Programme"

       ElseIf InStr(1, cel.Value, "QUALITY PLAN") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Quality Plan"

       ElseIf InStr(1, cel.Value, "QUALITY SYSTEM CERTIFICATE") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Quality System Certificate"

       ElseIf InStr(1, cel.Value, "RADIOGRAPHIC QUALIFICATION PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Radiographic Qualification Procedures"

       ElseIf InStr(1, cel.Value, "SURVEILLANCE PLANS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Surveillance Plans"

       ElseIf InStr(1, cel.Value, "WELDER QUALIFICATIONS PROGRAM & RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Welder Qualifications Program & Records"

       ElseIf InStr(1, cel.Value, "WELDING PROCEDURE QUALIFICATIONS RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Welding Procedure Qualifications Records (PQRs)"

       ElseIf InStr(1, cel.Value, "WELDING PROCEDURE SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Welding Procedure Specifications (WPS)"

       ElseIf InStr(1, cel.Value, "WELD MAP") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Weld Maps (Location/Type/NDE Req)"

       ElseIf InStr(1, cel.Value, "WELD REPAIR PROCEDURES") > 0 Then
            cel.Offset(0, -39).Value = "Quality Control"
            cel.Offset(0, -38).Value = "Weld Repair Procedures"

       ElseIf InStr(1, cel.Value, "BLOCK AND LEVEL DRAWING") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Block and Level Drawing"

       ElseIf InStr(1, cel.Value, "EQUIPMENT AND OUTLET LAYOUTS TELEPHONE, DATA, PA AND MATV SERVICES") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Equipment and Outlet Layouts Telephone, Data, PA and MATV Services"

       ElseIf InStr(1, cel.Value, "IDF RECORDS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "IDF Records"

       ElseIf InStr(1, cel.Value, "LISTS/BILL OF MATERIALS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Lists/Bill Of Materials"

       ElseIf InStr(1, cel.Value, "LOUDSPEAKER LOOP DIAGRAMS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Loudspeaker Loop Diagrams"

       ElseIf InStr(1, cel.Value, "OVERALL TELECOM EQUIPMENT AND CABLE ROUTING LAYOUT") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Overall Telecom Equipment and Cable Routing Layout"

       ElseIf InStr(1, cel.Value, "RADIO LICENCE") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Radio Licence"

       ElseIf InStr(1, cel.Value, "SYSTEM DRAWINGS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "System Drawings"

       ElseIf InStr(1, cel.Value, "TELECOM CABLE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Cable Schedule"

       ElseIf InStr(1, cel.Value, "TELECOM DATA SHEETS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Data Sheets"

       ElseIf InStr(1, cel.Value, "TELECOM EQUIPMENT SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Equipment Schedule"

       ElseIf InStr(1, cel.Value, "TELECOM INTERCONNECTION DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Interconnection Diagram"

       ElseIf InStr(1, cel.Value, "TELECOM LINE DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Line Diagram"

       ElseIf InStr(1, cel.Value, "TELECOM MANHOLE SCHEDULE") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Manhole Schedule"

       ElseIf InStr(1, cel.Value, "TELECOM MISCELLANEOUS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Miscellaneous"

       ElseIf InStr(1, cel.Value, "TELECOM SPECIFICATIONS") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Specifications"

       ElseIf InStr(1, cel.Value, "TELECOM WIRING DIAGRAM") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telecom Wiring Diagram"

       ElseIf InStr(1, cel.Value, "TELEMETRY ID LISTING") > 0 Then
            cel.Offset(0, -39).Value = "TeleCommunication"
            cel.Offset(0, -38).Value = "Telemetry ID Listing"
            
       ElseIf InStr(1, cel.Value, "MATERIAL REQUISITION") > 0 Then
            cel.Offset(0, -39).Value = "Procurement"
            cel.Offset(0, -38).Value = "Material Movement Request (MMR)"
                                
   End If
   
   Next cel
   
       SourceRange3.Offset(0, 47).ClearContents
       
End Sub

Private Sub GenerateUnits_Click()
Dim Tags As Variant
Dim TagUnit As Variant

Dim SrchRng As Range, cel As Range

Dim SourceRange As Range

On Error Resume Next

   Set SourceRange = Application.Selection
   Set SourceRange = Application.InputBox("Range:", "Select Tags", SourceRange.Address, Type:=8)
   
Err.Clear

On Error GoTo 0

   SourceRange.Offset(0, 2) = _
        "=IFERROR(LEFT(RC[-2],FIND(""-"",RC[-2],1)-1),LEFT(RC[-2],2))"
        

Set SrchRng = SourceRange

For Each cel In SrchRng
    If InStr(1, cel.Offset(0, 2).Value, "10") > 0 Then
        cel.Offset(0, 2).Value = "10-"

    ElseIf InStr(1, cel.Offset(0, 2).Value, "11") > 0 Then
        cel.Offset(0, 2).Value = "11-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "12") > 0 Then
        cel.Offset(0, 2).Value = "12-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "13") > 0 Then
        cel.Offset(0, 2).Value = "13-"

    ElseIf InStr(1, cel.Offset(0, 2).Value, "14") > 0 Then
        cel.Offset(0, 2).Value = "14-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "15") > 0 Then
        cel.Offset(0, 2).Value = "15-"
               
    ElseIf InStr(1, cel.Offset(0, 2).Value, "16") > 0 Then
        cel.Offset(0, 2).Value = "16-"
                             
    ElseIf InStr(1, cel.Offset(0, 2).Value, "19") > 0 Then
        cel.Offset(0, 2).Value = "19-"
    
    ElseIf InStr(1, cel.Offset(0, 2).Value, "52") > 0 Then
        cel.Offset(0, 2).Value = "52-"
        
    ElseIf InStr(1, cel.Offset(0, 2).Value, "45") > 0 Then
        cel.Offset(0, 2).Value = "45-"
     
    End If
Next cel
        
    For Each Tags In SourceRange.Offset(0, 2)
        If Tags = "14-" Then
             Tags.Offset(0, 5).Value = "Unit-14"
        
        ElseIf Tags = "13-" Then
             Tags.Offset(0, 5).Value = "Unit-13"
        ElseIf Tags = "11-" Then
             Tags.Offset(0, 5).Value = "Unit-11"
        ElseIf Tags = "12-" Then
             Tags.Offset(0, 5).Value = "Unit-12"
        ElseIf Tags = "15-" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "16-" Then
             Tags.Offset(0, 5).Value = "Unit-16"
        ElseIf Tags = "10-" Then
             Tags.Offset(0, 5).Value = "Unit-10"
        ElseIf Tags = "45-" Then
             Tags.Offset(0, 5).Value = "Unit-45"
        ElseIf Tags = "52-" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "83" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "19-" Then
             Tags.Offset(0, 5).Value = "Unit-19"
        ElseIf Tags = "84" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS1" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS2" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS3" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS4" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "SS5" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "41" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "42" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "43" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "44" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "452" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "47" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "48" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "49" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "900" Then
             Tags.Offset(0, 5).Value = "Unit-900"
        ElseIf Tags = "573" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "574" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "401" Then
             Tags.Offset(0, 5).Value = "Unit-573"
        ElseIf Tags = "605" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "015" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "059" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "17" Then
             Tags.Offset(0, 5).Value = "Unit-15"
        ElseIf Tags = "99" Then
             Tags.Offset(0, 5).Value = "Unit-99"
        ElseIf Tags = "" Then
             Tags.Offset(0, 5).Value = ""
        ElseIf Tags = "50" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "51" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "904" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "96" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "BRC" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "WH" Then
             Tags.Offset(0, 5).Value = "Unit-00"
        ElseIf Tags = "21" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "22" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "25" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "50" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "51" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "545" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "61" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "62" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "64" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "65" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "66" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "67" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "68" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "70" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "71" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "72" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "80" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "81" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "82" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "080" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        ElseIf Tags = "92" Then
             Tags.Offset(0, 5).Value = "Non-BAB"
        Else
            Tags.Offset(0, 5).Value = ""
    End If
    
    Next Tags
    
    For Each TagUnit In SourceRange.Offset(0, 7)
        If TagUnit = "Unit-10" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-11" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-12" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-13" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-14" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-15" Then
            TagUnit.Offset(0, -1).Value = "Utility"
        ElseIf TagUnit = "Unit-16" Then
            TagUnit.Offset(0, -1).Value = "Utility"
        ElseIf TagUnit = "Unit-19" Then
            TagUnit.Offset(0, -1).Value = "Process"
        ElseIf TagUnit = "Unit-00" Then
            TagUnit.Offset(0, -1).Value = "Common"
        ElseIf TagUnit = "Unit-99" Then
            TagUnit.Offset(0, -1).Value = "Common"
        ElseIf TagUnit = "Unit-45" Then
            TagUnit.Offset(0, -1).Value = "Pipelines"
        ElseIf TagUnit = "Unit-900" Then
            TagUnit.Offset(0, -1).Value = "Pipelines"
        ElseIf TagUnit = "Unit-573" Then
            TagUnit.Offset(0, -1).Value = "Process"

    End If
    
    Next TagUnit

    
    Range("C:C").ClearContents
    Range("D:D").ClearContents
    Range("E:E").ClearContents
    Range("F:F").ClearContents
End Sub

Private Sub GetFileFormat_Click()


Dim SourceRange2 As Range

On Error Resume Next
   
   Set SourceRange2 = Application.Selection
   Set SourceRange2 = Application.InputBox("Range:", "Selece Filenames: ", SourceRange2.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0
   
   SourceRange2.Offset(0, 29).Value = "=UPPER(RIGHT(RC[-29],4))"
   
   Application.ScreenUpdating = False
   
   For Each fileform In SourceRange2.Offset(0, 29)
   
    If fileform = ".PDF" Then
        fileform.Value = "PDF"
        
    ElseIf fileform = ".DWG" Then
        fileform.Value = "CAD"

    ElseIf fileform = ".DGN" Then
        fileform.Value = "CAD"
        
    ElseIf fileform = ".XLS" Then
        fileform.Value = "XLS"
        
     ElseIf fileform = "XLSX" Then
        fileform.Value = "XLS"
                      
     ElseIf fileform = ".DOC" Then
        fileform.Value = "DOC"
             
    Else
        fileform.Value = ""
        
    End If
    
    Next fileform
    
      Application.ScreenUpdating = True
    
     SourceRange2.Offset(0, 27).Value = "ADNOC Gas Processing"
     SourceRange2.Offset(0, 12).Value = "Pipeline Network"
   
      
End Sub

Private Sub Label2_Click()

End Sub
