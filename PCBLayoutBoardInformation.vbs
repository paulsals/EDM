' PCBLayoutBoardInformation.vb
' Script to retrieve and export board information from Expedition PCB
' This script collects various board properties and exports them to a JSON file

Option Explicit

' Main function that orchestrates the board information collection process
Sub Main()
    ' Get application and document objects
    Dim app
    Dim doc
    Dim jsonFilePath
    ' Removed type declarations for variables
    
    ' Create JSON data dictionary
    Dim boardInfo
    Set boardInfo = CreateObject("Scripting.Dictionary")
    
    ' Connect to application and get document
    Set app = GetApplication()
    Set doc = GetActiveDocument(app)
    
    ' Verify document is licensed
    If Not LicenseDocument(doc) Then
        MsgBox "Failed to license document for access", vbExclamation
        Exit Sub
    End If
    
    ' Get basic document properties
    boardInfo.Add "Name", GetDocumentName(doc)
    boardInfo.Add "BaseUnit", GetBaseUnit(doc)
    boardInfo.Add "LayerCount", GetLayerCount(doc)
    
    ' Get board outline coordinates
    Dim outlineCoordinates
    Set outlineCoordinates = GetOutlineCoordinates(doc)
    boardInfo.Add "MinX", outlineCoordinates("MinX")
    boardInfo.Add "MinY", outlineCoordinates("MinY")
    boardInfo.Add "MaxX", outlineCoordinates("MaxX")
    boardInfo.Add "MaxY", outlineCoordinates("MaxY")
    
    ' Get component counts
    boardInfo.Add "ComponentTotal", GetComponentTotal(doc)
    boardInfo.Add "TestPointTotal", GetTestPointTotal(doc)
    boardInfo.Add "TestPointTopTotal", GetTestPointTopTotal(doc)
    boardInfo.Add "TestPointBottomTotal", GetTestPointBottomTotal(doc)
    
    ' Get via and connection information
    Dim viasCollection
    Set viasCollection = GetViasCollection(doc)
    boardInfo.Add "ViaCount", GetViaCount(viasCollection)
    boardInfo.Add "ConnectionCountOption", GetConnectionCountOption(doc)
    
    ' Get nets information
    Dim netsCollection
    Set netsCollection = GetNetsCollection(doc)
    boardInfo.Add "NetCount", GetNetCount(netsCollection)
    
    ' Check for KANBAN cell
    boardInfo.Add "HasKanbanCell", CheckForKanbanCell(doc)
    
    ' Get hole information
    Dim holeInfo
    Set holeInfo = GetHoleInformation(doc)
    boardInfo.Add "SmallestDrillSize", holeInfo("SmallestDrillSize")
    boardInfo.Add "NonPlatedHoleCount", holeInfo("NonPlatedHoleCount")
    
    ' Calculate the total plated hole count
    Dim totalHoleCount
    Dim viaCount
    Dim nonPlatedHoleCount
    
    viaCount = boardInfo("ViaCount")
    nonPlatedHoleCount = holeInfo("NonPlatedHoleCount")
    totalHoleCount = GetTotalHoleCount(doc)
    
    boardInfo.Add "PlatedHoleCount", totalHoleCount - nonPlatedHoleCount - viaCount
    
    ' Write to JSON file
    jsonFilePath = WriteToJsonFile(boardInfo)
    
    MsgBox "Board information exported to: " & jsonFilePath, vbInformation

    ' Add logging to track variable population
    logMessage "Application object created."
    logMessage "Document object retrieved."
    logMessage "JSON file path: " & jsonFilePath
    logMessage "Board name: " & boardInfo("Name")
    logMessage "Base unit: " & boardInfo("BaseUnit")
    logMessage "Layer count: " & boardInfo("LayerCount")
    logMessage "Outline coordinates: MinX=" & outlineCoordinates("MinX") & ", MinY=" & outlineCoordinates("MinY") & ", MaxX=" & outlineCoordinates("MaxX") & ", MaxY=" & outlineCoordinates("MaxY")
    logMessage "Component total: " & boardInfo("ComponentTotal")
    logMessage "Test point total: " & boardInfo("TestPointTotal")
    logMessage "Via count: " & boardInfo("ViaCount")
    logMessage "Net count: " & boardInfo("NetCount")
    logMessage "Has Kanban cell: " & boardInfo("HasKanbanCell")
    logMessage "Smallest drill size: " & holeInfo("SmallestDrillSize")
    logMessage "Non-plated hole count: " & holeInfo("NonPlatedHoleCount")
    logMessage "Plated hole count: " & boardInfo("PlatedHoleCount")
End Sub

' Checks for the presence of a KANBAN cell in the document
Function CheckForKanbanCell(doc)
    Dim cells
    Dim cell
    
    Set cells = doc.Cells
    
    For Each cell In cells
        If UCase(cell.Name) = "KANBAN" Then
            CheckForKanbanCell = True
            Exit Function
        End If
    Next
    
    CheckForKanbanCell = False
End Function

' Gets the active document from the application
Function GetActiveDocument(app)
    Set GetActiveDocument = app.ActiveDocument
End Function

' Gets the Expedition PCB application object
Function GetApplication()
    Dim app
    On Error Resume Next
    Set app = GetObject(, "ExpPCB.Application")
    
    If app Is Nothing Then
        Set app = CreateObject("ExpPCB.Application")
    End If
    
    app.Visible = True
    Set GetApplication = app
End Function

' Gets the base unit of the document (IN or MM)
Function GetBaseUnit(doc)
    Dim baseUnit
    
    Select Case doc.BaseUnit
        Case 0
            baseUnit = "IN"
        Case 1
            baseUnit = "MM"
        Case Else
            baseUnit = "Unknown"
    End Select
    
    GetBaseUnit = baseUnit
End Function

' Gets the component total count
Function GetComponentTotal(doc)
    Dim components
    Set components = doc.Components
    GetComponentTotal = components.Count
End Function

' Gets the connection count option property
Function GetConnectionCountOption(doc)
    GetConnectionCountOption = doc.ConnectionCountOption
End Function

' Gets the document name
Function GetDocumentName(doc)
    GetDocumentName = doc.Name
End Function

' Gets hole information including smallest drill size and non-plated hole count
Function GetHoleInformation(doc)
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim holes
    Dim hole
    Dim smallestDrill
    Dim nonPlatedCount
    
    Set holes = doc.Holes
    smallestDrill = 9999.9 ' Large initial value
    nonPlatedCount = 0
    
    For Each hole In holes
        If hole.DrillSize < smallestDrill Then
            smallestDrill = hole.DrillSize
        End If
        
        If Not hole.Plated Then
            nonPlatedCount = nonPlatedCount + 1
        End If
    Next
    
    result.Add "SmallestDrillSize", smallestDrill
    result.Add "NonPlatedHoleCount", nonPlatedCount
    
    Set GetHoleInformation = result
End Function

' Gets the number of layers in the document
Function GetLayerCount(doc)
    Dim layers
    Set layers = doc.Layers
    GetLayerCount = layers.Count
End Function

' Gets the net count
Function GetNetCount(nets)
    GetNetCount = nets.Count
End Function

' Gets the nets collection
Function GetNetsCollection(doc)
    Set GetNetsCollection = doc.Nets
End Function

' Gets the board outline coordinates
Function GetOutlineCoordinates(doc)
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim outline
    Set outline = doc.Outline
    
    result.Add "MinX", outline.MinX
    result.Add "MinY", outline.MinY
    result.Add "MaxX", outline.MaxX
    result.Add "MaxY", outline.MaxY
    
    Set GetOutlineCoordinates = result
End Function

' Gets components with RefDes of TP on the bottom side
Function GetTestPointBottomTotal(doc)
    Dim components
    Dim component
    Dim count
    
    Set components = doc.Components
    count = 0
    
    For Each component In components
        If Left(component.RefDes, 2) = "TP" And component.Side = 1 Then ' Assuming 1 is bottom
            count = count + 1
        End If
    Next
    
    GetTestPointBottomTotal = count
End Function

' Gets the total count of test points (RefDes starting with TP)
Function GetTestPointTotal(doc)
    Dim components
    Dim component
    Dim count
    
    Set components = doc.Components
    count = 0
    
    For Each component In components
        If Left(component.RefDes, 2) = "TP" Then
            count = count + 1
        End If
    Next
    
    GetTestPointTotal = count
End Function

' Gets components with RefDes of TP on the top side
Function GetTestPointTopTotal(doc)
    Dim components
    Dim component
    Dim count
    
    Set components = doc.Components
    count = 0
    
    For Each component In components
        If Left(component.RefDes, 2) = "TP" And component.Side = 0 Then ' Assuming 0 is top
            count = count + 1
        End If
    Next
    
    GetTestPointTopTotal = count
End Function

' Gets the total hole count in the document
Function GetTotalHoleCount(doc)
    Dim holes
    Set holes = doc.Holes
    GetTotalHoleCount = holes.Count
End Function

' Gets the via count
Function GetViaCount(vias)
    GetViaCount = vias.Count
End Function

' Gets the vias collection
Function GetViasCollection(doc)
    Set GetViasCollection = doc.Vias
End Function

' Licenses the document for read access
Function LicenseDocument(doc)
    On Error Resume Next
    doc.License "Read"
    LicenseDocument = (Err.Number = 0)
End Function

' Writes the board information to a JSON file
Function WriteToJsonFile(boardInfo)
    Dim fso
    Dim jsonFile
    Dim filePath
    Dim json
    Dim key
    
    ' Create JSON string
    json = "{"
    For Each key In boardInfo.Keys
        ' Add quotes for string values
        If VarType(boardInfo(key)) = vbString Then
            json = json & """" & key & """: """ & boardInfo(key) & """, "
        Else
            json = json & """" & key & """: " & boardInfo(key) & ", "
        End If
    Next
    
    ' Remove the trailing comma and space
    If Len(json) > 1 Then
        json = Left(json, Len(json) - 2)
    End If
    
    json = json & "}"
    
    ' Write to file
    Set fso = CreateObject("Scripting.FileSystemObject")
    'filePath = fso.GetSpecialFolder(2) & "\BoardInfo_" & Format(Now, "yyyymmdd_hhnnss") & ".json"
    filePath = "C:\Scripts\JSON\test.json"
    
    Set jsonFile = fso.CreateTextFile(filePath, True)
    jsonFile.Write json
    jsonFile.Close
    
    WriteToJsonFile = filePath
End Function

' Logs a message to the troubleshooting file
Dim troubleshootingFilePath
troubleshootingFilePath = "c:\Scripts\JSON\troubleshooting.txt"

Sub logMessage(message)
    Dim fso, logFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set logFile = fso.OpenTextFile(troubleshootingFilePath, 8, True)
    logFile.WriteLine Now & " - " & message
    logFile.Close
End Sub