Option Explicit

Function GetCurrentPage() As Visio.page

    Debug.Print Application.ActivePage
    'MsgBox "Active Page: " + Application.ActivePage.Name
    
    Set GetCurrentPage = Application.ActivePage
    

End Function


Function CountShapesInPage(page As Visio.page) As Collection
    
    ''''''''''''''''''''''''
    ' Variable Definitions '
    ''''''''''''''''''''''''
    
    ' Define variables
    Dim totalShapeCount As Integer
    Dim pageTotaltotalShapeCount As Integer
    Dim shapesCollection As New Collection
    Dim shape As Visio.shape
    Dim regex_dataSpool As Object
    
    ' Regex variables
    Dim regex_dataFiber
    Dim regex_dataWallTV
    Dim regex_dataCoax
    Dim regex_WAP
    Dim regex_dataStdDrop
    Dim regex_dataWallPhone
    
    ' Count variables
    Dim dataSpool
    Dim dataFiber
    Dim dataWallTV
    Dim dataCoax
    Dim WAP
    Dim dataStdDrop
    Dim dataWallPhone
    
    ' Instantiate count variables with initial value of 0
    '
    '   dataSpool: integer
    '   dataFiber: integer
    '   dataWallTV: integer
    '   dataCoax: integer
    '   WAP: integer
    '   dataStdDrop: integer
    '   dataWallPhone: integer
    dataSpool = 0
    dataFiber = 0
    dataWallTV = 0
    dataCoax = 0
    WAP = 0
    dataStdDrop = 0
    dataWallPhone = 0
    
    ' Set total shape count variables
    pageTotaltotalShapeCount = page.Shapes.Count
    totalShapeCount = 0
    
    
    '''''''''
    ' REGEX '
    '''''''''
    
    ' Instantiate regex variables
    Set regex_dataSpool = New RegExp
    Set regex_dataFiber = New RegExp
    Set regex_dataWallTV = New RegExp
    Set regex_dataCoax = New RegExp
    Set regex_WAP = New RegExp
    Set regex_dataStdDrop = New RegExp
    Set regex_dataWallPhone = New RegExp
    
    ' Define regex patterns and settings
    regex_dataSpool.Pattern = "DataSpoolCeiling(\.\d*)?"
    regex_dataSpool.Global = True
    regex_dataSpool.IgnoreCase = True
    
    regex_dataFiber.Pattern = "DataFiber(\.\d*)?"
    regex_dataFiber.Global = True
    regex_dataFiber.IgnoreCase = True
    
    regex_dataWallTV.Pattern = "DataWallTV(\.\d*)?"
    regex_dataWallTV.Global = True
    regex_dataWallTV.IgnoreCase = True
    
    regex_dataCoax.Pattern = "DataCoax(\.\d*)?"
    regex_dataCoax.Global = True
    regex_dataCoax.IgnoreCase = True
    
    regex_WAP.Pattern = "WAP(\.\d*)?"
    regex_WAP.Global = True
    regex_WAP.IgnoreCase = True
    
    regex_dataStdDrop.Pattern = "DataStdDrop(\.\d*)?"
    regex_dataStdDrop.Global = True
    regex_dataStdDrop.IgnoreCase = True
    
    regex_dataWallPhone.Pattern = "DataWallPhone(\.\d*)?"
    regex_dataWallPhone.Global = True
    regex_dataWallPhone.IgnoreCase = True
    
    
    ''''''''
    ' Loop '
    ''''''''
    
    ' Loop through all shapes on current sheet to check if they are the correct stencil
    ' If they are, increment the respective type (i.e. data fiber matched, count = count + 1)
    For Each shape In page.Shapes
'    Debug.Print shape.Name
'    Debug.Print regex_dataSpool.Test(shape.Name)
    
    If regex_dataSpool.Test(shape.Name) Then
        dataSpool = dataSpool + 1
        Debug.Print "incremented data spool count " & CStr(dataSpool)
    
    ElseIf regex_dataFiber.Test(shape.Name) Then
        dataFiber = dataFiber + 1
        Debug.Print "incremented data fiber count " & CStr(dataFiber)
        
    ElseIf regex_dataWallTV.Test(shape.Name) Then
        dataWallTV = dataWallTV + 1
        Debug.Print "incremented data wall TV count " & CStr(dataWallTV)
        
    ElseIf regex_dataCoax.Test(shape.Name) Then
        dataCoax = dataCoax + 1
        Debug.Print "incremented data coax count " & CStr(dataCoax)
        
    ElseIf regex_WAP.Test(shape.Name) Then
        WAP = WAP + 1
        Debug.Print "incremented WAP count " & CStr(WAP)
        
    ElseIf regex_dataStdDrop.Test(shape.Name) Then
        dataStdDrop = dataStdDrop + 1
        Debug.Print "incremented dataStdDrop count " & CStr(dataStdDrop)
        
    ElseIf regex_dataWallPhone.Test(shape.Name) Then
        dataWallPhone = dataWallPhone + 1
        Debug.Print "incremented dataWallPhone count " & CStr(dataWallPhone)
    
    Else
        Debug.Print "Didn't find " & shape.Name
        
    End If
    
    Next shape
    
    
    
    '''''''''''''''''''''''''
    ' Final data processing '
    '''''''''''''''''''''''''
    
    ' Calculate total shapes found
    totalShapeCount = dataSpool + dataFiber + dataWallTV + dataCoax + WAP + dataStdDrop + dataWallPhone
    
    ' Construct a collection to be returned:
    '   dataSpool: integer
    '   dataFiber: integer
    '   dataWallTV: integer
    '   dataCoax: integer
    '   WAP: integer
    '   dataStdDrop: integer
    '   dataWallPhone: integer
    '   totalShapeCount: integer
    shapesCollection.Add Item:=CStr(dataSpool), Key:="dataSpool"
    shapesCollection.Add Item:=CStr(dataFiber), Key:="dataFiber"
    shapesCollection.Add Item:=CStr(dataWallTV), Key:="dataWallTV"
    shapesCollection.Add Item:=CStr(dataCoax), Key:="dataCoax"
    shapesCollection.Add Item:=CStr(WAP), Key:="WAP"
    shapesCollection.Add Item:=CStr(dataStdDrop), Key:="dataStdDrop"
    shapesCollection.Add Item:=CStr(dataWallPhone), Key:="dataWallPhone"
    shapesCollection.Add Item:=CStr(totalShapeCount), Key:="totalShapeCount"
    
    '''''''''''''''''
    ' DEBUG RESULTS '
    '''''''''''''''''
    
    'Debug.Print CStr(pageTotaltotalShapeCount) & " pageTotaltotalShapeCount"
    'Debug.Print shapesCollection.Item("totalShapeCount") & " total shapes found"
    'Debug.Print shapesCollection.Item("dataFiber") & " data fiber shapes found"
    
    Set CountShapesInPage = shapesCollection

End Function



Private Sub Document_DocumentOpened(ByVal doc As IVDocument)


    Call UpdateCountText
    
    'MsgBox "total shapes found" & shapesCollection.Item("totalShapesCount")
    

End Sub


' Updates values on screen => void
Public Sub UpdateCountText()

    ' Variables
    Dim currentPage As Visio.page
    
    Dim totalDropsText As String
    Dim totalDataSpoolText As String
    Dim totalDataFiberText As String
    Dim totalWallTVText As String
    Dim totalDataCoaxText As String
    Dim totalWAPText As String
    Dim totalDataStdDropText As String
    Dim totalDataWallPhoneText As String
    
    Dim totalDropsShape As String
    Dim totalDataSpoolShape As String
    Dim totalDataFiberShape As String
    Dim totalWallTVShape As String
    Dim totalDataCoaxShape As String
    Dim totalWAPShape As String
    Dim totalDataStdDropShape As String
    Dim totalDataWallPhoneShape As String
    
    ' Set shape names as this will differ from sheet to sheet
    totalDropsShape = "Sheet.1060"
    totalDataSpoolShape = "Sheet.2511"
    totalDataFiberShape = "Sheet.5975"
    totalWallTVShape = "Sheet.2519"
    totalDataCoaxShape = "Sheet.5976"
    totalWAPShape = "Sheet.2509"
    totalDataStdDropShape = "Sheet.6204"
    totalDataWallPhoneShape = "Sheet.6203"
    
    ' Get Current Page
    Set currentPage = GetCurrentPage
    
    ' Craft text for text boxes
    totalDropsText = "Total drops: " + CountShapesInPage(currentPage).Item("totalShapeCount")
    totalDataSpoolText = "Total data spool drops: " + CountShapesInPage(currentPage).Item("dataSpool")
    totalDataFiberText = "Total data fiber drops: " + CountShapesInPage(currentPage).Item("dataFiber")
    totalWallTVText = "Total Wall Mounted TV Locations: " + CountShapesInPage(currentPage).Item("dataWallTV")
    totalDataCoaxText = "Total coax drops: " + CountShapesInPage(currentPage).Item("dataCoax")
    totalWAPText = "Total WAPs: " + CountShapesInPage(currentPage).Item("WAP")
    totalDataStdDropText = "Total wall data drops: " + CountShapesInPage(currentPage).Item("dataStdDrop")
    totalDataWallPhoneText = "Total wall phone drops: " + CountShapesInPage(currentPage).Item("dataWallPhone")
    
    
    ' Assign shape count to text boxes on screen
    Dim shape
    
    For Each shape In currentPage.Shapes
        Select Case shape.Name
            Case totalDropsShape
                shape.Text = totalDropsText
                
            Case totalDataSpoolShape
                shape.Text = totalDataSpoolText
                
            Case totalDataFiberShape
                shape.Text = totalDataFiberText
                
            Case totalWallTVShape
                shape.Text = totalWallTVText
                
            Case totalDataCoaxShape
                shape.Text = totalDataCoaxText
                
            Case totalWAPShape
                shape.Text = totalWAPText
                
            Case totalDataStdDropShape
                shape.Text = totalDataStdDropText
            
            Case totalDataWallPhoneShape
                shape.Text = totalDataWallPhoneText
                
            Case Else
                Debug.Print "No match: " & shape.Name
                
        End Select
        
    Next

End Sub

Private Sub Document_DocumentSaved(ByVal doc As IVDocument)

    Call UpdateCountText

End Sub
