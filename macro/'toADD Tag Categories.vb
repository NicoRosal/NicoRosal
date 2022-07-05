'toADD Tag Categories

    ElseIf InStr(1, cel.Value, "TRV") > 0 Then
        cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
        cel.Offset(0, 3).Value = "RELIEF DEVICE"
        cel.Offset(0, 4).Value = "TEMPERATURE RELIEF VALVE"

    ElseIf InStr(1, cel.Value, "TRR") > 0 Then
        cel.Offset(0, 2).Value = "ELECTRICAL"
        cel.Offset(0, 3).Value = "MISCELLANOUS"
        cel.Offset(0, 4).Value = "CATHODIC PROTECTION RECTIFIER"

    ElseIf InStr(1, cel.Value, "PDIH") > 0 Then
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "PDIA") > 0 Then
        cel.Offset(0, 1).Value = "DIFFERENTIAL PRESSURE INDICATING ALARM"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "FIA") > 0 Then
        cel.Offset(0, 1).Value = "FLOW RATE ALARM"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "SVC") > 0 Then
        cel.Offset(0, 1).Value = "FLOW RATE ALARM"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "PANEL AND CONSOLE"

    ElseIf InStr(1, cel.Value, "ISO") > 0 Then
        cel.Offset(0, 1).Value = "ISOLATOR"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "WIRING"
        cel.Offset(0, 4).Value = "GALVANIC ISOLATOR"

    ElseIf InStr(1, cel.Value, "HOV") > 0 Then
        cel.Offset(0, 1).Value = "HAND OPERATED VALVE"    
        cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
        cel.Offset(0, 3).Value = "RELIEF DEVICE"
        cel.Offset(0, 4).Value = "PRESSURE RELIEF VALVE"

    ElseIf InStr(1, cel.Value, "PDH") > 0 Then
        cel.Offset(0, 1).Value = "Differential Pressure Alarm High"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "PDIH") > 0 Then
        cel.Offset(0, 1).Value = "Differential Pressure Alarm High"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "PDL") > 0 Then
        cel.Offset(0, 1).Value = "Differential Pressure Alarm Low"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "PDIL") > 0 Then
        cel.Offset(0, 1).Value = "Differential Pressure Alarm Low"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "PDIL") > 0 Then
        cel.Offset(0, 1).Value = "Differential Pressure Alarm Low"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"

    ElseIf InStr(1, cel.Value, "EDB") > 0 Then   
        cel.Offset(0, 2).Value = "ELECTRICAL"
        cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
        cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"

    ElseIf InStr(1, cel.Value, "PRD") > 0 Then
        cel.Offset(0, 1).Value = "PRESSURE RELIEF DAMPER"    
        cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
        cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
        cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"

    ElseIf InStr(1, cel.Value, "PDS") > 0 Then
        cel.Offset(0, 1).Value = "PRESSURE DIFFERENTIAL SWITCH"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "PRESSURE"
        cel.Offset(0, 4).Value = "PRESSURE SWITCH"

    ElseIf InStr(1, cel.Value, "DBBV") > 0 Then
        cel.Offset(0, 1).Value = "Double Block & Bleed Valve"    
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "VALVE"
        cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"

    ElseIf InStr(1, cel.Value, "MPGV") > 0 Then
        cel.Offset(0, 1).Value = "MULTIPORT GAUGE VALVE"    
        cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
        cel.Offset(0, 3).Value = "VALVE"
        cel.Offset(0, 4).Value = "GATE VALVE"

    ElseIf InStr(1, cel.Value, "-PB-") > 0 Then   
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "PUSH BUTTON"

    ElseIf InStr(1, cel.Value, "SSV") > 0 Then  
        cel.Offset(0, 1).Value = "SLAM SHUT VALVE"         
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "VALVE"
        cel.Offset(0, 4).Value = "TIGHT SHUT OFF VALVE"

    ElseIf InStr(1, cel.Value, "ANN") > 0 Then  
        cel.Offset(0, 1).Value = "Annunciator"         
        cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
        cel.Offset(0, 3).Value = "GENERAL"
        cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
