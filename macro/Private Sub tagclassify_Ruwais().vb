Private Sub tagclassify_Ruwais()


   Application.ScreenUpdating = False
   
   
   Err.Clear
    
        
    For Each cel In SourceRange3.Offset(0, 1)

        If InStr(1, cel.Value, Chr(34)) > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "PIPELINE"
            cel.Offset(0, 4).Value = "PIPERUN"
            
        ElseIf InStr(1, cel.Value, Chr(39)) > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "PIPELINE"
            cel.Offset(0, 4).Value = "PIPERUN"
            
        ElseIf InStr(1, cel.Value, "-XY-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "SOLENOID VALVE"

        ElseIf InStr(1, cel.Value, "-XY-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

        ElseIf InStr(1, cel.Value, "VDR") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "CHECK VALVE"

        ElseIf InStr(1, cel.Value, "VS") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "GATE VALVE"

        ElseIf InStr(1, cel.Value, "ME") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "MISCELLANEOUS EQUIPMENT"
            cel.Offset(0, 4).Value = "STEAM GENERATOR"

        ElseIf InStr(1, cel.Value, "ME") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "MISCELLANEOUS EQUIPMENT"
            cel.Offset(0, 4).Value = "STEAM GENERATOR"

        ElseIf InStr(1, cel.Value, "-EF-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "MISCELLANEOUS EQUIPMENT"
            cel.Offset(0, 4).Value = "STEAM GENERATOR"

        ElseIf InStr(1, cel.Value, "TK") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "STORAGE VESSELS/ TANKS"
            cel.Offset(0, 4).Value = "DOME-ROOF TANK"

        ElseIf InStr(1, cel.Value, "CV-M") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRIC MOTOR"
            cel.Offset(0, 4).Value = "INDUCTION MOTOR"

        ElseIf InStr(1, cel.Value, "CVM") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRIC MOTOR"
            cel.Offset(0, 4).Value = "INDUCTION MOTOR"

        ElseIf InStr(1, cel.Value, "GDM") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRIC MOTOR"
            cel.Offset(0, 4).Value = "INDUCTION MOTOR"

        ElseIf InStr(1, cel.Value, "CV") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "BELT CONVEYOR"

        ElseIf InStr(1, cel.Value, "MD") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"

        ElseIf InStr(1, cel.Value, "MFD") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"

        ElseIf InStr(1, cel.Value, "XJ") > 0 Then
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"

        ElseIf InStr(1, cel.Value, "STR") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID-SEPERATION EQUIPMENT"
            cel.Offset(0, 4).Value = "STRAINER"

        ElseIf InStr(1, cel.Value, "YA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "BE") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "INCINERATOR/ BURNER"
            cel.Offset(0, 4).Value = "ELEVATED-HYDROCARBONS-LIQUID BURNER"

        ElseIf InStr(1, cel.Value, "TML") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "VOLTAGE TRANSFORMER"

        ElseIf InStr(1, cel.Value, "-GD-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "PROCESS VESSELS"
            cel.Offset(0, 4).Value = "ELLIPTICAL-HEAD HORIZONTAL DRUM"

        ElseIf InStr(1, cel.Value, "SWG") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "HV SWITCHGEAR"

        ElseIf InStr(1, cel.Value, "EDH") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC DUCT HEATER"

        ElseIf InStr(1, cel.Value, "XAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "XAL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "TAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "TSH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "HS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"

        ElseIf InStr(1, cel.Value, "FACU") > 0 Then
            cel.Offset(0, 1).Value = "FRESH AIR CONDITIONING UNIT"
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"

        ElseIf InStr(1, cel.Value, "ACCU") > 0 Then
            cel.Offset(0, 1).Value = "AIR COOLED CONDENSING UNIT"
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"

        ElseIf InStr(1, cel.Value, "CWCP") > 0 Then
            cel.Offset(0, 1).Value = "CHILLED WATER CIRCULATION PUMP"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "IN-LINE CENTRIFUGAL PUMP"

        ElseIf InStr(1, cel.Value, "RACU") > 0 Then
            cel.Offset(0, 1).Value = "RECIRCULATION AIR-CON UNIT"
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"

        ElseIf InStr(1, cel.Value, "XSH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        ElseIf InStr(1, cel.Value, "SSL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"

        ElseIf InStr(1, cel.Value, "LED") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRIC LOAD"
            cel.Offset(0, 4).Value = "LIGHTING"
            
        ElseIf InStr(1, cel.Value, "FSL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PRESSURE SWITCH"

        ElseIf InStr(1, cel.Value, "CAV") > 0 Then
            cel.Offset(0, 1).Value = "CONSTANT AIR VOLUME TERMINAL"        
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"

        ElseIf InStr(1, cel.Value, "EBS") > 0 Then
            cel.Offset(0, 1).Value = "ETHERNET BACKBONE SWITCH"        
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"

        ElseIf InStr(1, cel.Value, "BSL") > 0 Then
            cel.Offset(0, 1).Value = "FLAME SCANNER"        
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"

        ElseIf InStr(1, cel.Value, "WIT") > 0 Then    
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"

        ElseIf InStr(1, cel.Value, "VFD") > 0 Then
            cel.Offset(0, 1).Value = "VOLUME FLOW DAMPER"        
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"

        ElseIf InStr(1, cel.Value, "VFD") > 0 Then      
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "DC UPS"

        ElseIf InStr(1, cel.Value, "THP") > 0 Then      
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "TACED HEATING PANEL"

        ElseIf InStr(1, cel.Value, "SEM") > 0 Then   
            cel.Offset(0, 1).Value = "VIBRATORY SCREEN MOTOR"            
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "EQUIPMENT DRIVER"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR"

        ElseIf InStr(1, cel.Value, "RPS") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "DC UPS"

        ElseIf InStr(1, cel.Value, "P-M") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "EQUIPMENT DRIVER"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR"

        ElseIf InStr(1, cel.Value, "P-M") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"

        ElseIf InStr(1, cel.Value, "IDB") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "POWER DISTRIBUTION BOARD"

        ElseIf InStr(1, cel.Value, "DGM") > 0 Then   
            cel.Offset(0, 1).Value = "DIVERTER GATE MOTOR"            
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "EQUIPMENT DRIVER"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR"

        ElseIf InStr(1, cel.Value, "PRS") > 0 Then   
            cel.Offset(0, 1).Value = "PIPE RACK SYSTEM"            
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "PIPE RACK"

        ElseIf InStr(1, cel.Value, "TC") > 0 Then   
            cel.Offset(0, 1).Value = "THERMOSTAT"            
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE GAGE"

        ElseIf InStr(1, cel.Value, "TE") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "THERMOCOUPLE TEMPERATURE ASSEMBLY"

        ElseIf InStr(1, cel.Value, "LT") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"

        ElseIf InStr(1, cel.Value, "LS") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL SWITCH"

        ElseIf InStr(1, cel.Value, "LI") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"

        ElseIf InStr(1, cel.Value, "HR") > 0 Then              
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "IN-LINE FITTING"
            cel.Offset(0, 4).Value = "CONNECTION, HOSE"

        ElseIf InStr(1, cel.Value, "HM") > 0 Then   
            cel.Offset(0, 1).Value = "TWO WAY HYDRANT WITH FIRE MONITOR"           
            cel.Offset(0, 2).Value = "HSE/ FIRE FIGHTING"
            cel.Offset(0, 3).Value = "FIRE FIGHTING ITEMS"
            cel.Offset(0, 4).Value = "HYDRANT POST(QUADRUPLE) WITH MONITOR"

        ElseIf InStr(1, cel.Value, "HB") > 0 Then              
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "IN-LINE FITTING"
            cel.Offset(0, 4).Value = "CONNECTION, HOSE"

        ElseIf InStr(1, cel.Value, "GS") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"

        ElseIf InStr(1, cel.Value, "FT") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"

        ElseIf InStr(1, cel.Value, "FO") > 0 Then              
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"

        ElseIf InStr(1, cel.Value, "EF") > 0 Then              
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "AXIAL FAN"

       ElseIf InStr(1, cel.Value, "SE") > 0 Then    
            cel.Offset(0, 1).Value = "VIBRATORY SCREEN"             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID-SEPERATION EQUIPMENT"
            cel.Offset(0, 4).Value = "SINGLE-PRODUCT SCREEN"

       ElseIf InStr(1, cel.Value, "DG") > 0 Then 
            cel.Offset(0, 1).Value = "GRANULATION DRUM"             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "PROCESS VESSELS"
            cel.Offset(0, 4).Value = "ELLIPTICAL-HEAD HORIZONTAL DRUM"

       ElseIf InStr(1, cel.Value, "IY") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "VIBRATING LEVEL SWITCH"

'BATCH 2

       ElseIf InStr(1, cel.Value, "FAVU") > 0 Then             
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"

       ElseIf InStr(1, cel.Value, "DPIT") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"

       ElseIf InStr(1, cel.Value, "PDIT") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"

       ElseIf InStr(1, cel.Value, "GZCO") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"

       ElseIf InStr(1, cel.Value, "ACPC") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "WATER-CHILLING EVAPORATOR"

       ElseIf InStr(1, cel.Value, "EVR") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "DISCONNECT ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "OVERLOAD RELAY"

       ElseIf InStr(1, cel.Value, "EWC") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "WATER-CHILLING EVAPORATOR"

       ElseIf InStr(1, cel.Value, "HWG") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL HEATER"

       ElseIf InStr(1, cel.Value, "PLC") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "SOFTWARE FUNCTION"

       ElseIf InStr(1, cel.Value, "SAC") > 0 Then             
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"

       ElseIf InStr(1, cel.Value, "TIT") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE TRANSMITTER"

       ElseIf InStr(1, cel.Value, "PGA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

       ElseIf InStr(1, cel.Value, "UEA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "VAH") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "ESD") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "WQI") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "HIA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "HXF") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

      ElseIf InStr(1, cel.Value, "GEC") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

      ElseIf InStr(1, cel.Value, "GEV") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

      ElseIf InStr(1, cel.Value, "GES") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

     ElseIf InStr(1, cel.Value, "AEC") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL CABINET"

      ElseIf InStr(1, cel.Value, "GZA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "RLP") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"

      ElseIf InStr(1, cel.Value, "PLP") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"

      ElseIf InStr(1, cel.Value, "MVP") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "MOTERIZED VALVE PANEL"

      ElseIf InStr(1, cel.Value, "TBP") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"

      ElseIf InStr(1, cel.Value, "FAL") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "EZA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "UEA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "WAL") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "MAF") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

      ElseIf InStr(1, cel.Value, "SO2") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "TOXIC GAS DETECTOR"

      ElseIf InStr(1, cel.Value, "ESB") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "HV SWITCHGEAR"

   ElseIf InStr(1, cel.Value, "PFC") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BATTERY BANK"

   ElseIf InStr(1, cel.Value, "FA") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "AXIAL FAN"

      ElseIf InStr(1, cel.Value, "HY") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
    
      ElseIf InStr(1, cel.Value, "NL") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "NY") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

      ElseIf InStr(1, cel.Value, "VI") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

     ElseIf InStr(1, cel.Value, "ZA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"

     ElseIf InStr(1, cel.Value, "WE") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"

     ElseIf InStr(1, cel.Value, "WI") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "WEIGHT TRANSMITTER"

     ElseIf InStr(1, cel.Value, "FS") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

     ElseIf InStr(1, cel.Value, "EC") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL CABINET"

    ElseIf InStr(1, cel.Value, "YT") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "POSITIVE DISPLACEMENT FLOW TRANSMITTER"

    ElseIf InStr(1, cel.Value, "-FQ-") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID-SEPERATION EQUIPMENT"
            cel.Offset(0, 4).Value = "STRAINER"

     ElseIf InStr(1, cel.Value, "DS") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

     ElseIf InStr(1, cel.Value, "RA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"

     ElseIf InStr(1, cel.Value, "BD") > 0 Then             
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BUS DUCT"

     ElseIf InStr(1, cel.Value, "CT") > 0 Then             
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL HEATER"

     ElseIf InStr(1, cel.Value, "TL") > 0 Then             
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"

          
        End If
   
    Next cel
    
End Sub

Private Sub tagclassify_Ruwais_2()


   Application.ScreenUpdating = False
   
   
   Err.Clear
    
        
    For Each cel In SourceRange3.Offset(0, 1)

     ElseIf InStr(1, cel.Value, "HXA") > 0 Then             
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

        End If
   
    Next cel
    
End Sub