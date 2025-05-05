' PCBLayoutBoardInformation.vb
' Script to retrieve and export board information from Expedition PCB
' This script collects various board properties and exports them to a JSON file

Option Explicit

' Main function that orchestrates the board information collection process
Sub Main()
    ' Get application and document objects
    Dim app As Object
    Dim doc As Object
    Dim jsonFilePath As String
    
    ' Create JSON data dictionary
    Dim boardInfo As Object
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
    Dim outlineCoordinates As Object
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
    Dim viasCollection As Object
    Set viasCollection = GetViasCollection(doc)
    boardInfo.Add "ViaCount", GetViaCount(viasCollection)
    boardInfo.Add "ConnectionCountOption", GetConnectionCountOption(doc)
    
    ' Get nets information
    Dim netsCollection As Object
    Set netsCollection = GetNetsCollection(doc)
    boardInfo.Add "NetCount", GetNetCount(netsCollection)
    
    ' Check for KANBAN cell
    boardInfo.Add "HasKanbanCell", CheckForKanbanCell(doc)
    
    ' Get hole information
    Dim holeInfo As Object
    Set holeInfo = GetHoleInformation(doc)
    boardInfo.Add "SmallestDrillSize", holeInfo("SmallestDrillSize")
    boardInfo.Add "NonPlatedHoleCount", holeInfo("NonPlatedHoleCount")
    
    ' Calculate the total plated hole count
    Dim totalHoleCount As Long
    Dim viaCount As Long
    Dim nonPlatedHoleCount As Long
    
    viaCount = boardInfo("ViaCount")
    nonPlatedHoleCount = holeInfo("NonPlatedHoleCount")
    totalHoleCount = GetTotalHoleCount(doc)
    
    boardInfo.Add "PlatedHoleCount", totalHoleCount - nonPlatedHoleCount - viaCount
    
    ' Write to JSON file
    jsonFilePath = WriteToJsonFile(boardInfo)
    
    MsgBox "Board information exported to: " & jsonFilePath, vbInformation
End Sub

' Checks for the presence of a KANBAN cell in the document
Function CheckForKanbanCell(doc As Object) As Boolean
    Dim cells As Object
    Dim cell As Object
    
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
Function GetActiveDocument(app As Object) As Object
    Set GetActiveDocument = app.ActiveDocument
End Function

' Gets the Expedition PCB application object
Function GetApplication() As Object
    Dim app As Object
    On Error Resume Next
    Set app = GetObject(, "ExpPCB.Application")
    
    If app Is Nothing Then
        Set app = CreateObject("ExpPCB.Application")
    End If
    
    app.Visible = True
    Set GetApplication = app
End Function

' Gets the base unit of the document (IN or MM)
Function GetBaseUnit(doc As Object) As String
    Dim baseUnit As String
    
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
Function GetComponentTotal(doc As Object) As Long
    Dim components As Object
    Set components = doc.Components
    GetComponentTotal = components.Count
End Function

' Gets the connection count option property
Function GetConnectionCountOption(doc As Object) As Long
    GetConnectionCountOption = doc.ConnectionCountOption
End Function

' Gets the document name
Function GetDocumentName(doc As Object) As String
    GetDocumentName = doc.Name
End Function

' Gets hole information including smallest drill size and non-plated hole count
Function GetHoleInformation(doc As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim holes As Object
    Dim hole As Object
    Dim smallestDrill As Double
    Dim nonPlatedCount As Long
    
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
Function GetLayerCount(doc As Object) As Long
    Dim layers As Object
    Set layers = doc.Layers
    GetLayerCount = layers.Count
End Function

' Gets the net count
Function GetNetCount(nets As Object) As Long
    GetNetCount = nets.Count
End Function

' Gets the nets collection
Function GetNetsCollection(doc As Object) As Object
    Set GetNetsCollection = doc.Nets
End Function

' Gets the board outline coordinates
Function GetOutlineCoordinates(doc As Object) As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim outline As Object
    Set outline = doc.Outline
    
    result.Add "MinX", outline.MinX
    result.Add "MinY", outline.MinY
    result.Add "MaxX", outline.MaxX
    result.Add "MaxY", outline.MaxY
    
    Set GetOutlineCoordinates = result
End Function

' Gets components with RefDes of TP on the bottom side
Function GetTestPointBottomTotal(doc As Object) As Long
    Dim components As Object
    Dim component As Object
    Dim count As Long
    
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
Function GetTestPointTotal(doc As Object) As Long
    Dim components As Object
    Dim component As Object
    Dim count As Long
    
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
Function GetTestPointTopTotal(doc As Object) As Long
    Dim components As Object
    Dim component As Object
    Dim count As Long
    
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
Function GetTotalHoleCount(doc As Object) As Long
    Dim holes As Object
    Set holes = doc.Holes
    GetTotalHoleCount = holes.Count
End Function

' Gets the via count
Function GetViaCount(vias As Object) As Long
    GetViaCount = vias.Count
End Function

' Gets the vias collection
Function GetViasCollection(doc As Object) As Object
    Set GetViasCollection = doc.Vias
End Function

' Licenses the document for read access
Function LicenseDocument(doc As Object) As Boolean
    On Error Resume Next
    doc.License "Read"
    LicenseDocument = (Err.Number = 0)
End Function

' Writes the board information to a JSON file
Function WriteToJsonFile(boardInfo As Object) As String
    Dim fso As Object
    Dim jsonFile As Object
    Dim filePath As String
    Dim json As String
    Dim key As Variant
    
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
    filePath = fso.GetSpecialFolder(2) & "\BoardInfo_" & Format(Now, "yyyymmdd_hhnnss") & ".json"
    
    Set jsonFile = fso.CreateTextFile(filePath, True)
    jsonFile.Write json
    jsonFile.Close
    
    WriteToJsonFile = filePath
End Function