Dim SourceRange3 As Range, cel As Range
Sub Delete_Content()
Dim mbResult As Integer
mbResult = MsgBox("These changes cannot be undone. Would you like to proceed?", _
 vbYesNoCancel)

Select Case mbResult
    Case vbYes
                With ActiveWorkbook
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
        End With
    Case vbNo

    Case vbCancel
        Exit Sub

End Select
End Sub
Sub Classify_Tags()
On Error Resume Next

   Set SourceRange3 = Application.Selection
   Set SourceRange3 = Application.InputBox("Range:", "Select Tags: ", SourceRange3.Address, Type:=8)
   
   Err.Clear

On Error GoTo 0

   Application.ScreenUpdating = False
    
    SourceRange3.Offset(0, 1).Value = "=UPPER(RC[-1])"
    SourceRange3.Offset(0, 6).Value = "PIPELINE NETWORK"
        
    For Each cel In SourceRange3.Offset(0, 1)
            
        If InStr(1, cel.Value, Chr(34)) > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "PIPELINE"
            cel.Offset(0, 4).Value = "PIPERUN"
            
        ElseIf InStr(1, cel.Value, Chr(39)) > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "PIPELINE"
            cel.Offset(0, 4).Value = "PIPERUN"
            
        ElseIf InStr(1, cel.Value, "MOV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR OPERATED VALVE"
            
        ElseIf InStr(1, cel.Value, "ROV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR OPERATED VALVE"
            
        ElseIf InStr(1, cel.Value, "LVP") > 0 Then
            cel.Offset(0, 1).Value = "LOW VOLTAGE POWER CABLE"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
            
        ElseIf InStr(1, cel.Value, "-PC-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
            
        ElseIf InStr(1, cel.Value, "PG") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE GAUGE"
            
        ElseIf InStr(1, cel.Value, "PI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE GAUGE"
            
        ElseIf InStr(1, cel.Value, "ZSO") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION SWITCH"
            
        ElseIf InStr(1, cel.Value, "ZSC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION SWITCH"
            
        ElseIf InStr(1, cel.Value, "ZS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION SWITCH"
            
        ElseIf InStr(1, cel.Value, "ZL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "ZI") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "PCV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "-EE-SP") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-EE-LP") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"
            
        ElseIf InStr(1, cel.Value, "-EE-WO") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "SOCKET OUTLET"
            
        ElseIf InStr(1, cel.Value, "-EE-CP") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "CATHODIC PROTECTION RECTIFIER"
            
        ElseIf InStr(1, cel.Value, "-EE-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "CONTROL CABLE"
            
        ElseIf InStr(1, cel.Value, "ZT") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "JB") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "F-JD") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "F-JC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "FJD") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "FJC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "-CC-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "CONTROL CABLE"
            
        ElseIf InStr(1, cel.Value, "GDF") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-GD-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "AIT") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "ANALYZER TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "GDT") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "GDH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
        
        ElseIf InStr(1, cel.Value, "FDI") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-HD-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "TI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-TG-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE GAGE"
            
        ElseIf InStr(1, cel.Value, "-TG-") > 0 Then
            cel.Offset(0, 1).Value = "Temperature - Indicator"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE GAGE"
            
        ElseIf InStr(1, cel.Value, "TW-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "THERMOWELL"
            
        ElseIf InStr(1, cel.Value, "RO-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "RESTRICTION ORIFICE"
            
        ElseIf InStr(1, cel.Value, "-XA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-RHA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MRHC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MZY") > 0 Then
            cel.Offset(0, 1).Value = "MOV CLOSE PERMISSIVE"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MRHO") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MRXO") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "RHSO") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "RHSC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MRXC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "XS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-HS-") > 0 Then
            cel.Offset(0, 1).Value = "Hand (Operated) - Switch"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-SD-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-UV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "QVS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "SOV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "SOLENOID VALVE"
            
        ElseIf InStr(1, cel.Value, "FE-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW ELEMENT"
            
        ElseIf InStr(1, cel.Value, "FG-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW GAGE"
            
        ElseIf InStr(1, cel.Value, "AE-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "ANALYZER SAMPLE TAKE OFF PROBE"
            
        ElseIf InStr(1, cel.Value, "XUY") > 0 Then
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "COMM") > 0 Then
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "FNG") > 0 Then
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "FFA") > 0 Then
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-FC-") > 0 Then
            cel.Offset(0, 1).Value = "FLOW COMP"
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-PRT") > 0 Then
            cel.Offset(0, 1).Value = "LASER PRINTER"
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-EL-") > 0 Then
            cel.Offset(0, 1).Value = "RTU PANEL"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-FA-") > 0 Then
            cel.Offset(0, 1).Value = "RTU PANEL"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "TK") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "STORAGE VESSELS/ TANKS"
            cel.Offset(0, 4).Value = "CONE-ROOF TANK"
            
        ElseIf InStr(1, cel.Value, "-LI-") > 0 Then
            cel.Offset(0, 1).Value = "LEVEL TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-PS-") > 0 Then
            cel.Offset(0, 1).Value = "PRESSURE SWITCH"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE SWITCH"
            
        ElseIf InStr(1, cel.Value, "IFH") > 0 Then
            cel.Offset(0, 1).Value = "INDOOR FIRE HYDRANT"
            cel.Offset(0, 2).Value = "HSE/ FIRE FIGHTING"
            cel.Offset(0, 3).Value = "FIRE FIGHTING ITEMS"
            cel.Offset(0, 4).Value = "HYDRANT POST(DOUBLE)"
            
        ElseIf InStr(1, cel.Value, "OFH") > 0 Then
            cel.Offset(0, 1).Value = "OUTDOOR FIRE HYDRANT"
            cel.Offset(0, 2).Value = "HSE/ FIRE FIGHTING"
            cel.Offset(0, 3).Value = "FIRE FIGHTING ITEMS"
            cel.Offset(0, 4).Value = "HYDRANT POST(DOUBLE)"
            
        ElseIf InStr(1, cel.Value, "FPH") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "VOLTAGE TRANSFORMER"
            
        ElseIf InStr(1, cel.Value, "PSV") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "RELIEF DEVICE"
            cel.Offset(0, 4).Value = "PRESSURE RELIEF VALVE"
            
        ElseIf InStr(1, cel.Value, "-XV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ON/OFF VALVE"
            
        ElseIf InStr(1, cel.Value, "-PT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "ABSOLUTE PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-ME-") > 0 Then
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "EQUIPMENT FOUNDATION"
            
        ElseIf InStr(1, cel.Value, "-SP-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "SPECIALITY PIPING ITEMS"
            
        ElseIf InStr(1, cel.Value, "-FV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "FLOW CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "-PV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "-PIC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-LT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-PY-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "RELAY PANEL"
            
        ElseIf InStr(1, cel.Value, "-TE-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "THERMOCOUPLE TEMPERATURE ASSEMBLY"
            
        ElseIf InStr(1, cel.Value, "-EV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR OPERATED VALVE"
            
        ElseIf InStr(1, cel.Value, "-FT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-LAH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-LAHH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PAL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PALL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-FI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-LAL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-LALL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PAH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PAHH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "HSO-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "HSC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-HS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "VE-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "VIBRATION PROBE"
            
        ElseIf InStr(1, cel.Value, "-SC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "SPEED TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "PDT") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "HCV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "LIC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "XASS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-CH-") > 0 Then
            cel.Offset(0, 1).Value = "CHRLORINE PACKAGE VESSEL"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "PROCESS VESSELS"
            cel.Offset(0, 4).Value = "ELLIPTICAL-HEAD VERTICAL DRUM"
            
        ElseIf InStr(1, cel.Value, "XASS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-MRS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "GSP") > 0 Then
            cel.Offset(0, 1).Value = "CO2 SUPPRESSION PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "ICP") > 0 Then
            cel.Offset(0, 1).Value = "INERGEN CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "PM-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "IN-LINE CENTRIFUGAL PUMP"
            
        ElseIf InStr(1, cel.Value, "-DB-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "-MDB-") > 0 Then
            cel.Offset(0, 2).Value = "MAIN DISTRIBUTION BOARD"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "-ICB-") > 0 Then
            cel.Offset(0, 2).Value = "LIGHTING DISTRIBUTION BOARD"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "-MCC-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "RP-") > 0 Then
            cel.Offset(0, 1).Value = "RELAY PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "RELAY PANEL"
            
        ElseIf InStr(1, cel.Value, "-HVAC-") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR MOVING DEVICES AND COMPONENTS"
            
        ElseIf InStr(1, cel.Value, "WED-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "PDAH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "PSLL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "YAHH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "FAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "XL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "ISHH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "IAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "FY") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW ELEMENT"
            
        ElseIf InStr(1, cel.Value, "PDAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "PDI") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "HV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ON/OFF VALVE"
            
        ElseIf InStr(1, cel.Value, "PSHH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "ZAL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "ZAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PAHH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XY") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "LOCAL ESD SWITCH"
            
        ElseIf InStr(1, cel.Value, "-HC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-SU-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "STORAGE VESSELS/ TANKS"
            cel.Offset(0, 4).Value = "SUMP"
            
        ElseIf InStr(1, cel.Value, "-YI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-YE-") > 0 Then
            cel.Offset(0, 1).Value = "TEMP ELEMENT"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "THERMOCOUPLE TEMPERATURE ASSEMBLY"
            
        ElseIf InStr(1, cel.Value, "-FIC-") > 0 Then
            cel.Offset(0, 1).Value = "FLOW - CONTROLLER, INDICATING"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "VARIABLE AREA FLOW INDICATOR"
            
        ElseIf InStr(1, cel.Value, "TDC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-II-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "BATTERY CHARGER"
            
        ElseIf InStr(1, cel.Value, "-PA-") > 0 Then
            cel.Offset(0, 1).Value = "PRESSURE ALARM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE SWITCH"
            
        ElseIf InStr(1, cel.Value, "-PIC-") > 0 Then
            cel.Offset(0, 1).Value = "PRESSURE INDICATING CONTROLLER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE GAUGE"
            
        ElseIf InStr(1, cel.Value, "-PI") > 0 Then
            cel.Offset(0, 1).Value = "PRESSURE INDICATOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-LG-") > 0 Then
            cel.Offset(0, 1).Value = "LEVEL GAUGE"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL GAUGE GLASS"
            
        ElseIf InStr(1, cel.Value, "-HY-") > 0 Then
            cel.Offset(0, 1).Value = "Hand (Operated) - Relay/Positioner/Computing Function"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "I/P CONVERTOR"
            
        ElseIf InStr(1, cel.Value, "-PIT-") > 0 Then
            cel.Offset(0, 1).Value = "PRESSURE INDICATING TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "PRESSURE TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-TAH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-EA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "-HIS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-PDIT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-FAM-") > 0 Then
            cel.Offset(0, 1).Value = "FAN MOTOR"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "EQUIPMENT DRIVER"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR OPERATED VALVE"
            
        ElseIf InStr(1, cel.Value, "-HRSG-") > 0 Then
            cel.Offset(0, 1).Value = "Heat Recovery  Steam Generator"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "MISCELLANEOUS EQUIPMENT"
            cel.Offset(0, 4).Value = "STEAM GENERATOR"
            
        ElseIf InStr(1, cel.Value, "-AGM-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "EQUIPMENT DRIVER"
            cel.Offset(0, 4).Value = "ELECTRIC MOTOR"
            
        ElseIf InStr(1, cel.Value, "-AG-") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "MIXING AND BLENDING EQUIPMENT"
            cel.Offset(0, 4).Value = "AXIAL TURBINE MIXER"
            
        ElseIf InStr(1, cel.Value, "PDY") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "ELP") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "EMERGENCY LIGHTING PANEL"
            
        ElseIf InStr(1, cel.Value, "-DI-") > 0 Then
            cel.Offset(0, 1).Value = "DENSITY INDICATOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "ANALYZER TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-DT-") > 0 Then
            cel.Offset(0, 1).Value = "DENSITY TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "ANALYZER TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-TDO-") > 0 Then
            cel.Offset(0, 1).Value = "TEMP DETECTOR OPEN"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "HEAT DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-TD-") > 0 Then
            cel.Offset(0, 1).Value = "TEMP DETECTOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "HEAT DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-EAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-EAL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-TIT-") > 0 Then
            cel.Offset(0, 1).Value = "TEMPERATURE INDICATING TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-FIT-") > 0 Then
            cel.Offset(0, 1).Value = "FLOW INDICATING TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-FUY-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-V-101") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "SCRAPER FEEDER"
            
        ElseIf InStr(1, cel.Value, "-V-102") > 0 Then
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "SCRAPER CONVEYOR"
            
        ElseIf InStr(1, cel.Value, "-EI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-FAU-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-FQI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-FP-") > 0 Then
            cel.Offset(0, 1).Value = "Flow Rate Integrate Summation Indicating"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "-MTI") > 0 Then
            cel.Offset(0, 1).Value = "Maintenance Terminal"
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-EAU-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-POI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-BV-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "BALL VALVE"
            
        ElseIf InStr(1, cel.Value, "-NV-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "NEEDLE VALVE"
            
        ElseIf InStr(1, cel.Value, "-GV-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "GATE VALVE"
            
        ElseIf InStr(1, cel.Value, "-NRV-") > 0 Then
            cel.Offset(0, 1).Value = "NON RETURN VALVE"
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "CHECK VALVE"
            
        ElseIf InStr(1, cel.Value, "-PRV-") > 0 Then
            cel.Offset(0, 1).Value = "RETURN VALVE"
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "CHECK VALVE"
            
        ElseIf InStr(1, cel.Value, "-VB-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "BALL VALVE"
            
        ElseIf InStr(1, cel.Value, "-VC-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "CHECK VALVE"
            
        ElseIf InStr(1, cel.Value, "-VG-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "GATE VALVE"
            
        ElseIf InStr(1, cel.Value, "-VD-") > 0 Then
            cel.Offset(0, 1).Value = "DOUBLE BLOCK AND BLEED VALVE"
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "BALL VALVE"
            
        ElseIf InStr(1, cel.Value, "-VW-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "BUTTERFLY VALVE"
            
        ElseIf InStr(1, cel.Value, "-VP-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PLUG VALVE"
            
        ElseIf InStr(1, cel.Value, "-VL-") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "GLOBE VALVE"
            
        ElseIf InStr(1, cel.Value, "XZA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "XZV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-FO-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW ORIFICE PLATE"
            
        ElseIf InStr(1, cel.Value, "-XIA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XSI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-AT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-AI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XCV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ON/OFF VALVE"
            
        ElseIf InStr(1, cel.Value, "-TT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "ROY") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-HI-") > 0 Then
            cel.Offset(0, 2).Value = "HAND INDICATOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-RP-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "RELAY PANEL"
            
        ElseIf InStr(1, cel.Value, "RHPB") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "RELAY PANEL"
            
        ElseIf InStr(1, cel.Value, "RPB") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "RELAY PANEL"
            
        ElseIf InStr(1, cel.Value, "KV") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "BALL VALVE"
            
        ElseIf InStr(1, cel.Value, "LV") > 0 Then
            cel.Offset(0, 1).Value = "Level - Valve, Control"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "PRESSURE CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "LSH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "LSL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "LAH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "LAL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "XE") > 0 Then
            cel.Offset(0, 1).Value = "IGNITOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "DISPLACEMENT PROBE"
            
        ElseIf InStr(1, cel.Value, "UZ") > 0 Then
            cel.Offset(0, 1).Value = "POSITION INDICATION"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "DISPLACEMENT PROBE"
            
        ElseIf InStr(1, cel.Value, "LIT") > 0 Then
            cel.Offset(0, 1).Value = "LEVEL INDICATING TRANSMITTER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "PSL") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "PSH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-UA") > 0 Then
            cel.Offset(0, 1).Value = "LOGIC ALARM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-SDH") > 0 Then
            cel.Offset(0, 1).Value = "SMOKE DETECT ALARM HIGH"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "HDH") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "MHPB") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "PHPB") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "XSC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-YA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "LY") > 0 Then
            cel.Offset(0, 1).Value = "Level - Relay/Positioner/Computing Function"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "VOLUME MEASUREMENT EQUIPMENT"
            
        ElseIf InStr(1, cel.Value, "AH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "AHH-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "AL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "ALL-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-TV-") > 0 Then
            cel.Offset(0, 1).Value = "Temperature - Valve, Control"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "ON/OFF Valve"
            
        ElseIf InStr(1, cel.Value, "YS") > 0 Then
            cel.Offset(0, 1).Value = "VIBRATION SWITCH"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "PDV") > 0 Then
            cel.Offset(0, 1).Value = "Pressure Differential - Valve, Control"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "Valve"
            cel.Offset(0, 4).Value = "Pressure Control Valve"
            
        ElseIf InStr(1, cel.Value, "LIR") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "LEVEL"
            cel.Offset(0, 4).Value = "LEVEL TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "FIR") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW TRANSMITTER (OTHER)"
            
        ElseIf InStr(1, cel.Value, "TIC") > 0 Then
            cel.Offset(0, 1).Value = "TEMP INDICATING CONTROLLER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "FCV") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "FLOW CONTROL VALVE"
            
        ElseIf InStr(1, cel.Value, "PSF-") > 0 Then
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "PIPE SUPPORT FOUNDATION"
            
        ElseIf InStr(1, cel.Value, "KGD") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "KMC") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "BEACON"
            
        ElseIf InStr(1, cel.Value, "KFD") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "BEACON"
            
        ElseIf InStr(1, cel.Value, "-MX") > 0 Then
            cel.Offset(0, 1).Value = "MANUAL RELEASE"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-CS-") > 0 Then
            cel.Offset(0, 21).Value = "DELETE"
            cel.Offset(0, 7).Value = "DELETE"
            cel.Offset(0, 2).Value = "DELETE"
            
        ElseIf InStr(1, cel.Value, "-CP") > 0 Then
            cel.Offset(0, 1).Value = "CABLE"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
            
        ElseIf InStr(1, cel.Value, "-IJ") > 0 Then
            cel.Offset(0, 1).Value = "ISOLATION JOINT"
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "SPECIALITY PIPING ITEMS"
            cel.Offset(0, 4).Value = "JOINT BUTT WELDED"
            
        ElseIf InStr(1, cel.Value, "LF") > 0 Then
            cel.Offset(0, 1).Value = "LIGHT FITTING"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRIC LOAD"
            cel.Offset(0, 4).Value = "LIGHTING"
            
        ElseIf InStr(1, cel.Value, "HIC") > 0 Then
            cel.Offset(0, 1).Value = "HAND INDICATING CONTROLLER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "EP") > 0 Then
            cel.Offset(0, 1).Value = "EARTH PIT"
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "TRENCH"
            
        ElseIf InStr(1, cel.Value, "-SO-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "SOCKET OUTLET"
            
        ElseIf InStr(1, cel.Value, "-WO-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "WELDING OUTLETS"
            
        ElseIf InStr(1, cel.Value, "-ET-") > 0 Then
            cel.Offset(0, 1).Value = "ETHERNET SWITCH"
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-IT-") > 0 Then
            cel.Offset(0, 1).Value = "CURRENT - TRANSMITTER"
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
            
        ElseIf InStr(1, cel.Value, "-XS-") > 0 Then
            cel.Offset(0, 1).Value = "Safety Acting Switch"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "LOCAL ESD SWITCH"
            
        ElseIf InStr(1, cel.Value, "-XSO") > 0 Then
            cel.Offset(0, 1).Value = "Safety Acting Switch Open"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "LOCAL ESD SWITCH"
        
        ElseIf InStr(1, cel.Value, "-XSC") > 0 Then
            cel.Offset(0, 1).Value = "Safety Acting Switch Close"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "LOCAL ESD SWITCH"
            
        ElseIf InStr(1, cel.Value, "-FA-") > 0 Then
            cel.Offset(0, 1).Value = "FIRE DETECTION ALARM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-FL-") > 0 Then
            cel.Offset(0, 1).Value = "FILTER"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID-SEPERATION EQUIPMENT"
            cel.Offset(0, 4).Value = "GAS FILTER"
            
        ElseIf InStr(1, cel.Value, "-ZA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PAI-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-AP-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ALARM PANEL"
            
        ElseIf InStr(1, cel.Value, "-ES-") > 0 Then
            cel.Offset(0, 1).Value = "EMERGENCY SHUTDOWN"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-ESI-") > 0 Then
            cel.Offset(0, 1).Value = "EMERGENCY SHUTDOWN"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-ESH-") > 0 Then
            cel.Offset(0, 1).Value = "EMERGENCY SHUTDOWN HIGH"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-PD-") > 0 Then
            cel.Offset(0, 1).Value = "DIFF PRESSURE RANGE HIGH/LOW"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-FXA-") > 0 Then
            cel.Offset(0, 1).Value = "FAULT ALARM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-UPS-") > 0 Then
            cel.Offset(0, 1).Value = "UNINTERRUPTABLE POWER SUPPLY"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "AC UPS"
            
        ElseIf InStr(1, cel.Value, "-XIA") > 0 Then
            cel.Offset(0, 1).Value = "ALARM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-XCP") > 0 Then
            cel.Offset(0, 1).Value = "CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-XRR") > 0 Then
            cel.Offset(0, 1).Value = "HEAT DETECTOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "HEAT DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-TSH") > 0 Then
            cel.Offset(0, 1).Value = "THERMOSTAT"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE GAGE"
            
        ElseIf InStr(1, cel.Value, "-TIA-") > 0 Then
            cel.Offset(0, 1).Value = "Temperature - Indicator"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "TEMPERATURE GAGE"
            
        ElseIf InStr(1, cel.Value, "-TSL") > 0 Then
            cel.Offset(0, 1).Value = "LOW TEMPERATURE SIGNAL"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-LDB") > 0 Then
            cel.Offset(0, 1).Value = "LIGHTING DISTRIBUTION BOARD"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "PDB") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "POWER DISTRIBUTION BOARD"
            
        ElseIf InStr(1, cel.Value, "-LP") > 0 Then
            cel.Offset(0, 1).Value = "LIGHTING PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"
            
        ElseIf InStr(1, cel.Value, "-IR-") > 0 Then
            cel.Offset(0, 1).Value = "INTERFACE RELAY"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "DISCONNECT ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "PROTECTION RELAY"
            
        ElseIf InStr(1, cel.Value, "-TY-") > 0 Then
            cel.Offset(0, 1).Value = "Temperature - Relay/Positioner/Computing Function"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "TEMPERATURE"
            cel.Offset(0, 4).Value = "RESISTANCE TEMPERATURE ASSEMBLY"
            
        ElseIf InStr(1, cel.Value, "-GA-") > 0 Then
            cel.Offset(0, 1).Value = "Temperature - Relay/Positioner/Computing Function"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-SPP-") > 0 Then
            cel.Offset(0, 1).Value = "Solar Panel and Passive Shelter"
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "SHELTER FOUNDATION"
            
        ElseIf InStr(1, cel.Value, "-EG-") > 0 Then
            cel.Offset(0, 1).Value = "Egress Gate Deactivate"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-BA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "BEACON"
            
        ElseIf InStr(1, cel.Value, "-MA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "MANUAL CALL POINT"
            
        ElseIf InStr(1, cel.Value, "-MCP-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "MANUAL CALL POINT"
            
        ElseIf InStr(1, cel.Value, "-MC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "MANUAL CALL POINT"
            
        ElseIf InStr(1, cel.Value, "-RHS-") > 0 Then
            cel.Offset(0, 1).Value = "ROV Test Hand Switch"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-MHS-") > 0 Then
            cel.Offset(0, 1).Value = "Hand Switch Stop"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-MHSC-") > 0 Then
            cel.Offset(0, 1).Value = "Hand Switch Close"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-MHSO-") > 0 Then
            cel.Offset(0, 1).Value = "Hand Switch Open"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "CONTROL SWITCH"
            
        ElseIf InStr(1, cel.Value, "-HA-") > 0 Then
            cel.Offset(0, 1).Value = "Horns"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-BVS-") > 0 Then
            cel.Offset(0, 1).Value = "BLOCK VALVE STATION"
            cel.Offset(0, 2).Value = "CIVIL AND STRUCTURE"
            cel.Offset(0, 3).Value = "CIVIL ELEMENTS"
            cel.Offset(0, 4).Value = "EQUIPMENT STRUCTURE FOUNDATION"
            
        ElseIf InStr(1, cel.Value, "-SLP-") > 0 Then
            cel.Offset(0, 1).Value = "SOLAR POWERED PANELS"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-EDPV-") > 0 Then
            cel.Offset(0, 1).Value = "EMERGENCY DIFF PRESSURE VALVE"
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "RELIEF DEVICE"
            cel.Offset(0, 4).Value = "PRESSURE RELIEF VALVE"
            
        ElseIf InStr(1, cel.Value, "-ESDV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "EMERGENCY SHUTDOWN VALVE"
            
        ElseIf InStr(1, cel.Value, "-SDV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "EMERGENCY SHUTDOWN VALVE"
            
        ElseIf InStr(1, cel.Value, "-MZA-") > 0 Then
            cel.Offset(0, 1).Value = "OPEN POSITION INDICATION"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-PDIT-") > 0 Then
            cel.Offset(0, 1).Value = "DIFFERENTIAL PRESSURE INDICATOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-ES-") > 0 Then
            cel.Offset(0, 1).Value = "SOLAR POWER SYSTEM"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-ESL-") > 0 Then
            cel.Offset(0, 1).Value = "SOLAR POWER SYSTEM"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-ESH-") > 0 Then
            cel.Offset(0, 1).Value = "SOLAR POWER SYSTEM"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-QVT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "TOXIC GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-QGT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "FLAMMABLE GAS DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-MAC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "MANUAL CALL POINT"
            
        ElseIf InStr(1, cel.Value, "-RSV-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "VALVE"
            cel.Offset(0, 4).Value = "SOLENOID VALVE"
            
        ElseIf InStr(1, cel.Value, "-DPT-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "PRESSURE"
            cel.Offset(0, 4).Value = "DIFFERENTIAL PRESSURE TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-FB-") > 0 Then
            cel.Offset(0, 1).Value = "FIRE BELL"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
            
        ElseIf InStr(1, cel.Value, "-IC-") > 0 Then
            cel.Offset(0, 1).Value = "SMOKE DETECTOR BELL"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
            
        ElseIf InStr(1, cel.Value, "-FIQ-") > 0 Then
            cel.Offset(0, 1).Value = "FLOW INDICATING CONTROLLER"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "FLOW"
            cel.Offset(0, 4).Value = "VARIABLE AREA FLOW INDICATOR"
            
        ElseIf InStr(1, cel.Value, "-FJA-") > 0 Then
            cel.Offset(0, 1).Value = "F&G JUNCTION BOX"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "-FJD-") > 0 Then
            cel.Offset(0, 1).Value = "F&G JUNCTION BOX"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
            
        ElseIf InStr(1, cel.Value, "-PP-") > 0 Then
            cel.Offset(0, 1).Value = "SMALL POWER PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "SMALL POWER PANEL"
            
        ElseIf InStr(1, cel.Value, "-XT-") > 0 Then
            cel.Offset(0, 1).Value = "Miscellaneous - Transmitter"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "POSITION TRANSMITTER"
            
        ElseIf InStr(1, cel.Value, "-PHC-") > 0 Then
            cel.Offset(0, 1).Value = "PHOTO CELL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "LIGHTING PANEL"
            
        ElseIf InStr(1, cel.Value, "-BMS-") > 0 Then
            cel.Offset(0, 1).Value = "BATTERY MONITORING SYSTEM CABINET"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "SYSTEM CABINET"
            
        ElseIf InStr(1, cel.Value, "-BAT-") > 0 Then
            cel.Offset(0, 1).Value = "BATTERY MONITORING SYSTEM CABINET"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BATTERY BANK"
            
        ElseIf InStr(1, cel.Value, "-BAT-") > 0 Then
            cel.Offset(0, 1).Value = "BATTERY MONITORING SYSTEM CABINET"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BATTERY BANK"
            
        ElseIf InStr(1, cel.Value, "-CBB-") > 0 Then
            cel.Offset(0, 1).Value = "BATTERY CIRCUIT BREAKER"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "DISCONNECT ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "CIRCUIT BREAKER"
            
        ElseIf InStr(1, cel.Value, "-BT-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BATTERY BANK"
            
        ElseIf InStr(1, cel.Value, "-ISL-") > 0 Then
            cel.Offset(0, 1).Value = "BATTERY ISOLATION MCCB BOX"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "MISCELLANOUS"
            cel.Offset(0, 4).Value = "BATTERY BANK"
            
        ElseIf InStr(1, cel.Value, "-HCP-") > 0 Then
            cel.Offset(0, 1).Value = "HVAC CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-FGP-") > 0 Then
            cel.Offset(0, 1).Value = "F&G CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-FCP-") > 0 Then
            cel.Offset(0, 1).Value = "F&G CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-ACP-") > 0 Then
            cel.Offset(0, 1).Value = "AHU STARTER CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
    
        ElseIf InStr(1, cel.Value, "-BCP-") > 0 Then
            cel.Offset(0, 1).Value = "BLEED FAN FILTER STARTER CP"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-ECP-") > 0 Then
            cel.Offset(0, 1).Value = "EXHAUST FAN STARTER CP"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "-HSSD") > 0 Then
            cel.Offset(0, 1).Value = "HSSD PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            

        End If
   
    Next cel
    
    Call tagclassify
    Call CommandButton2_Click
    
    
    SourceRange3.Offset(0, 1).ClearContents
        
End Sub



Private Sub CommandButton2_Click()

       
       For Each DT In SourceRange3.Offset(0, 3)

    If DT = "INSTRUMENT AND CONTROL" Then
        DT.Offset(0, 6).Value = "Instrumentation"

    ElseIf DT = "MECHANICAL" Then
        DT.Offset(0, 6).Value = "Mechanical"
        
    ElseIf DT = "ELECTRICAL" Then
        DT.Offset(0, 6).Value = "Electrical"
        
    ElseIf DT = "PIPING AND PIPELINE" Then
        DT.Offset(0, 6).Value = "Piping"
        
    ElseIf DT = "MISCELLANEOUS" Then
        DT.Offset(0, 6).Value = "TeleCommunication"
        
    ElseIf DT = "HSE/ FIRE FIGHTING" Then
        DT.Offset(0, 6).Value = "HSE"
        
    ElseIf DT = "HVAC EQUIPMENT" Then
        DT.Offset(0, 6).Value = "HVAC"
        
    ElseIf DT = "CIVIL AND STRUCTURE" Then
        DT.Offset(0, 6).Value = "Civil & Structural"
        
    ElseIf DT = "DELETE" Then
        DT.Offset(0, 6).Value = "DELETE"
        DT.Offset(0, 19).Value = "DELETE"
             
    End If
    
        Next DT
        
       
       For Each ST In SourceRange3.Offset(0, 5)
       
    If ST = "PIPERUN" Then
        ST.Offset(0, 4).Value = "Pipeline"
        
    End If
    
        Next ST
End Sub

Private Sub tagclassify()


   Application.ScreenUpdating = False
   
   
   Err.Clear
    
        
    For Each cel In SourceRange3.Offset(0, 1)
            
         If InStr(1, cel.Value, "-FGS") > 0 Then
            cel.Offset(0, 1).Value = "F&G SYSTEM"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "INERGEN") > 0 Then
            cel.Offset(0, 1).Value = "INERGEN GAS PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "MIMIC") > 0 Then
            cel.Offset(0, 1).Value = "MIMIC PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "FMP") > 0 Then
            cel.Offset(0, 1).Value = "FLOW METERING PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
            
        ElseIf InStr(1, cel.Value, "AHU") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"
            
        ElseIf InStr(1, cel.Value, "-VA-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "VIBRATION PROBE"
            
        ElseIf InStr(1, cel.Value, "-CU-") > 0 Then
            cel.Offset(0, 1).Value = "CONDENSER UNIT"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "JET-EVAPORATIVE CONDENSER"
            
        ElseIf InStr(1, cel.Value, "-HTR-") > 0 Then
            cel.Offset(0, 1).Value = "ELEC HEATER"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "HEAT TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "ELECTRICAL HEATER"
            
        ElseIf InStr(1, cel.Value, "-EF-") > 0 Then
            cel.Offset(0, 1).Value = "EXHAUST FAN"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "AXIAL FAN"
            
        ElseIf InStr(1, cel.Value, "-EXF-") > 0 Then
            cel.Offset(0, 1).Value = "EXHAUST FAN"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "FLUID TRANSFER EQUIPMENT"
            cel.Offset(0, 4).Value = "AXIAL FAN"
            
        ElseIf InStr(1, cel.Value, "-BF-") > 0 Then
            cel.Offset(0, 1).Value = "FAN FILTER"
            cel.Offset(0, 2).Value = "MECHANICAL"
            cel.Offset(0, 3).Value = "SOLID-SEPERATION EQUIPMENT"
            cel.Offset(0, 4).Value = "BAG FILTER"
            
        ElseIf InStr(1, cel.Value, "-XIC-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
            
        ElseIf InStr(1, cel.Value, "-XIF-") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
            
        ElseIf InStr(1, cel.Value, "-LCP-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
            
        ElseIf InStr(1, cel.Value, "-DCS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "SYSTEM CABINET"
            
        ElseIf InStr(1, cel.Value, "-IPS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "SYSTEM CABINET"
            
        ElseIf InStr(1, cel.Value, "-TCS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "SYSTEM CABINET"
            
        ElseIf InStr(1, cel.Value, "-NWK-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "NETWORK CABINET"
            
        ElseIf InStr(1, cel.Value, "-TCP-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "NETWORK CABINET"
            
        ElseIf InStr(1, cel.Value, "-IMC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "MARSHALLING CABINET"
            
        ElseIf InStr(1, cel.Value, "-VMS-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "MARSHALLING CABINET"
            
        ElseIf InStr(1, cel.Value, "-DMC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "MARSHALLING CABINET"
            
        ElseIf InStr(1, cel.Value, "-FMC-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "CABINET"
            cel.Offset(0, 4).Value = "MARSHALLING CABINET"
            
            
        ElseIf InStr(1, cel.Value, "PACU") > 0 Then
            cel.Offset(0, 2).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 3).Value = "HVAC EQUIPMENT"
            cel.Offset(0, 4).Value = "HVAC AIR HANDLING UNIT"
            
        ElseIf InStr(1, cel.Value, "-SV-01") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
            
        ElseIf InStr(1, cel.Value, "-SV-02") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-SV-1") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
            
        ElseIf InStr(1, cel.Value, "-SV-2") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-SV-3") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-SV-03") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-SV1") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
            
        ElseIf InStr(1, cel.Value, "-SV2") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-SV3") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "-TP-") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
			
        ElseIf InStr(1, cel.Value, "CRS") > 0 Then
            cel.Offset(0, 2).Value = "DELETE"
                                                       
        ElseIf InStr(1, cel.Value, "-FD-") > 0 Then
            cel.Offset(0, 1).Value = "FIRE DETECTOR"		
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
			
        ElseIf InStr(1, cel.Value, "RDF") > 0 Then
            cel.Offset(0, 1).Value = "FIRE DETECTOR"		
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
			
        ElseIf InStr(1, cel.Value, "XGA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
			
        ElseIf InStr(1, cel.Value, "ZBS") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
			
        ElseIf InStr(1, cel.Value, "RXA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "COMMUNICATION SIGNAL"
			
        ElseIf InStr(1, cel.Value, "CCTV") > 0 Then
            cel.Offset(0, 1).Value = "CLOSED CIRCUIT TELEVISION"		
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
			
        ElseIf InStr(1, cel.Value, "TEL") > 0 Then
            cel.Offset(0, 1).Value = "CLOSED CIRCUIT TELEVISION"		
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
			
        ElseIf InStr(1, cel.Value, "CTV") > 0 Then		
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
			
        ElseIf InStr(1, cel.Value, "CWT") > 0 Then		
            cel.Offset(0, 2).Value = "MISCELLANEOUS"
            cel.Offset(0, 3).Value = "OTHERS"
            cel.Offset(0, 4).Value = "TELECOM DEVICES"
			
        ElseIf InStr(1, cel.Value, "FOC") > 0 Then
            cel.Offset(0, 1).Value = "FIBER OPTIC CABLE"		
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
			
        ElseIf InStr(1, cel.Value, "-EB-") > 0 Then
            cel.Offset(0, 1).Value = "ELECTRICAL EARTH BUSBAR (EB)"		
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "EARTHING CABLE"
			
        ElseIf InStr(1, cel.Value, "-FOPP-") > 0 Then
            cel.Offset(0, 1).Value = "FIBER OPTIC PATCH PANEL"		
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
			
        ElseIf InStr(1, cel.Value, "-IA-") > 0 Then
            cel.Offset(0, 1).Value = "SMOKE DETECTORS IN FIELD"		
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "SMOKE (FLAME) DETECTOR"
			
        ElseIf InStr(1, cel.Value, "-ACK-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
			
        ElseIf InStr(1, cel.Value, "-RST-") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "PANEL AND CONSOLE"
			
        ElseIf InStr(1, cel.Value, "PWR") > 0 Then
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
			
        ElseIf InStr(1, cel.Value, "FPC") > 0 Then
            cel.Offset(0, 1).Value = "FO PATCH CORD"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "ELECTRICAL CABLE"
            cel.Offset(0, 4).Value = "POWER CABLE"
			
        ElseIf InStr(1, cel.Value, "F-JA") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "GENERAL"
            cel.Offset(0, 4).Value = "INSTRUMENT JUNCTION BOX"
			
        ElseIf InStr(1, cel.Value, "TSV") > 0 Then
            cel.Offset(0, 2).Value = "PIPING AND PIPELINE"
            cel.Offset(0, 3).Value = "RELIEF DEVICE"
            cel.Offset(0, 4).Value = "TEMPERATURE RELIEF VALVE"
			
        ElseIf InStr(1, cel.Value, "GX") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "TOXIC GAS DETECTOR"
		
        ElseIf InStr(1, cel.Value, "QVI") > 0 Then
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "TOXIC GAS DETECTOR"
			
        ElseIf InStr(1, cel.Value, "RTD") > 0 Then
            cel.Offset(0, 1).Value = "RESISTANT TEMPERATURE DETECTOR"
            cel.Offset(0, 2).Value = "INSTRUMENT AND CONTROL"
            cel.Offset(0, 3).Value = "ANALYZER"
            cel.Offset(0, 4).Value = "HEAT DETECTOR"
            
        ElseIf InStr(1, cel.Value, "ACUPS") > 0 Then
            cel.Offset(0, 1).Value = "UNINTERRUPTABLE POWER SUPPLY"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "CONVERTING ELECTRICAL EQUIPMENT"
            cel.Offset(0, 4).Value = "AC UPS"
			
        ElseIf InStr(1, cel.Value, "FACP") > 0 Then
            cel.Offset(0, 1).Value = "FIRE ALARM CONTROL PANEL"
            cel.Offset(0, 2).Value = "ELECTRICAL"
            cel.Offset(0, 3).Value = "POWER DISTRIBUTION EQUIPMENT"
            cel.Offset(0, 4).Value = "CONTROL PANEL"
			
        End If
   
    Next cel
    
End Sub
Sub Del_Content()
Dim mbResult As Integer
mbResult = MsgBox("These changes cannot be undone. Would you like to proceed?", _
 vbYesNoCancel)

Select Case mbResult
    Case vbYes
                With ActiveWorkbook

    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A1").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("A2").Select
        End With
    Case vbNo

    Case vbCancel
        Exit Sub
        
End Select

End Sub
Sub DeleteRefDoc()

Dim mbResult As Integer
mbResult = MsgBox("These changes cannot be undone. Would you like to proceed?", _
 vbYesNoCancel)

Select Case mbResult
    Case vbYes
                With ActiveWorkbook
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
        End With
    Case vbNo

    Case vbCancel
        Exit Sub
        
End Select

End Sub
Sub Del_Automation()

Dim mbResult As Integer
mbResult = MsgBox("These changes cannot be undone. Would you like to proceed?", _
 vbYesNoCancel)

Select Case mbResult
    Case vbYes

    With ActiveWorkbook
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
        End With
    Case vbNo

    Case vbCancel
        Exit Sub
        
End Select
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
End Sub
