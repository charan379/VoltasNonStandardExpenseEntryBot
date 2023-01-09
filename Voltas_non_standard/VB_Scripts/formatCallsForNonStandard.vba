Sub VoltasCallsFormaterExcelNonStandard()
'Formats the CSV file from Vcare site into a ready made working format
'Delete not required columns
Dim a As Long, w As Long, vDELCOLs As Variant, vCOLNDX As Variant
vDELCOLs = Array("Symptom", "SR Value", "SC", "SC Band", "Threshold Value", "Total Value", "Manager", "Survey", "Survey Date", "UPBG Zone", "Deallocate Reason", "Row Id", "Commission Msg", "Created By", "TCR#", "TCR Date", "Mood of the customer", "Contact #", "TAT", "TAT Band", "Contact", "Fee Amount", "Program #", "EC", "Capacity", "Created by Division", "Product Group", "Calling from Number", "Type", "SAP Contract #", "Contract Type", "Agreement", "Address", "Cancel Reason", "Customer Comments", "Remarks", "Escalation", "Severity", "VIP", "Mobile Update", "Gas Charge Req Flag", "Audit Type", "Audit Date", "Purchased From Type", "Part Required Flag", "House #", "Building", "Road", "State", "Closure Code", "Purchased From", "Purchased From Free", "Last Modified By", "RT", "DT", "Attend time", "Appointment Date", "Serial# Source", "Split Serial# Source", "Serial Source Updated", "Split Serial Source Updated", "NPS Score", "Email Add", "Purchase Date")
With ThisWorkbook
    For w = 1 To .Worksheets.Count
	'With ActiveSheet.UsedRange  'Use This For ActiveSheet
        With Worksheets(w)
            For a = LBound(vDELCOLs) To UBound(vDELCOLs)
                vCOLNDX = Application.Match(vDELCOLs(a), .Rows(1), 0)
                If Not IsError(vCOLNDX) Then
                    .Columns(vCOLNDX).EntireColumn.Delete
                End If
            Next a
        End With
    Next w
End With

'Delete not required columns end

'Delete not required columns 2nd array

Dim a_b As Long, w_b As Long, vDELCOLs_b As Variant, vCOLNDX_b As Variant
vDELCOLs_b = Array("Invalid Code Remarks", "Organization", "External SR No", "Activation Key", "Promo Code", "Last Visit Date", "Tech Id", "FLS", "Closure Code Status", "ReOpen Count", "WTA/PR SR#", "WTA SR Status", "WTA Email Status","Ereceipt Status","Purchased From Code","External Ticket Id","SR Correction Flag","Model #","Age","FollowUp Count","Registered Phone","Status","Sub Status","Model","Serial #","Serial #(Split)","Key Account (S)","Key Account (P)","Owner","Open Date","Description","Alternate Phone","Account Type","Assign Date","Audit Status","Penalty Order#")
With ThisWorkbook
    For w_b = 1 To .Worksheets.Count
	'With ActiveSheet.UsedRange  'Use This For ActiveSheet
        With Worksheets(w_b)
            For a_b = LBound(vDELCOLs_b) To UBound(vDELCOLs_b)
                vCOLNDX_b = Application.Match(vDELCOLs_b(a_b), .Rows(1), 0)
                If Not IsError(vCOLNDX_b) Then
                    .Columns(vCOLNDX_b).EntireColumn.Delete
                End If
            Next a_b
        End With
    Next w_b
End With

'Delete not required columns end  2nd array

'Find SR # column
    Dim xRg_SR As Range
    Dim xRgUni_SR As Range
    Dim xAddress_SR As String
    Dim xStr_SR As String
    On Error Resume Next
    xStr_SR = "SR #"
    Set xRg_SR = ActiveSheet.UsedRange.Find(xStr_SR, , xlValues, xlWhole, , , True)
    If Not xRg_SR Is Nothing Then
        xAddress_SR = xRg_SR.Address
        Do
            Set xRg_SR = ActiveSheet.UsedRange.FindNext(xRg_SR)
            If xRgUni_SR Is Nothing Then
                Set xRgUni_SR = xRg_SR
            Else
                Set xRgUni_SR = Application.Union(xRgUni_SR, xRg_SR)
            End If
        Loop While (Not xRg_SR Is Nothing) And (xRg_SR.Address <> xAddress_SR)
    End If
    xRgUni_SR.EntireColumn.Activate
     With xRgUni_SR.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find SR # column End
    
'Find Service_Agent_Code # column
    Dim xRg_Service_Agent_Code As Range
    Dim xRgUni_Service_Agent_Code As Range
    Dim xAddress_Service_Agent_Code As String
    Dim xStr_Service_Agent_Code As String
    On Error Resume Next
    xStr_Service_Agent_Code = "Service Agent #"
    Set xRg_Service_Agent_Code = ActiveSheet.UsedRange.Find(xStr_Service_Agent_Code, , xlValues, xlWhole, , , True)
    If Not xRg_Service_Agent_Code Is Nothing Then
        xAddress_Service_Agent_Code = xRg_Service_Agent_Code.Address
        Do
            Set xRg_Service_Agent_Code = ActiveSheet.UsedRange.FindNext(xRg_Service_Agent_Code)
            If xRgUni_Service_Agent_Code Is Nothing Then
                Set xRgUni_Service_Agent_Code = xRg_Service_Agent_Code
            Else
                Set xRgUni_Service_Agent_Code = Application.Union(xRgUni_Service_Agent_Code, xRg_Service_Agent_Code)
            End If
        Loop While (Not xRg_Service_Agent_Code Is Nothing) And (xRg_Service_Agent_Code.Address <> xAddress_Service_Agent_Code)
    End If
    xRgUni_Service_Agent_Code.EntireColumn.Activate
     With xRgUni_Service_Agent_Code.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Service_Agent_Code # column End
    
        
'Find Call_Type # column
    Dim xRg_Call_Type As Range
    Dim xRgUni_Call_Type As Range
    Dim xAddress_Call_Type As String
    Dim xStr_Call_Type As String
    On Error Resume Next
    xStr_Call_Type = "Call Type"
    Set xRg_Call_Type = ActiveSheet.UsedRange.Find(xStr_Call_Type, , xlValues, xlWhole, , , True)
    If Not xRg_Call_Type Is Nothing Then
        xAddress_Call_Type = xRg_Call_Type.Address
        Do
            Set xRg_Call_Type = ActiveSheet.UsedRange.FindNext(xRg_Call_Type)
            If xRgUni_Call_Type Is Nothing Then
                Set xRgUni_Call_Type = xRg_Call_Type
            Else
                Set xRgUni_Call_Type = Application.Union(xRgUni_Call_Type, xRg_Call_Type)
            End If
        Loop While (Not xRg_Call_Type Is Nothing) And (xRg_Call_Type.Address <> xAddress_Call_Type)
    End If
    xRgUni_Call_Type.EntireColumn.Activate
     With xRgUni_Call_Type.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Call_Type # column END


'Find Account # column
    Dim xRg_Account As Range
    Dim xRgUni_Account As Range
    Dim xAddress_Account As String
    Dim xStr_Account As String
    On Error Resume Next
    xStr_Account = "Account"
    Set xRg_Account = ActiveSheet.UsedRange.Find(xStr_Account, , xlValues, xlWhole, , , True)
    If Not xRg_Account Is Nothing Then
        xAddress_Account = xRg_Account.Address
        Do
            Set xRg_Account = ActiveSheet.UsedRange.FindNext(xRg_Account)
            If xRgUni_Account Is Nothing Then
                Set xRgUni_Account = xRg_Account
            Else
                Set xRgUni_Account = Application.Union(xRgUni_Account, xRg_Account)
            End If
        Loop While (Not xRg_Account Is Nothing) And (xRg_Account.Address <> xAddress_Account)
    End If
    xRgUni_Account.EntireColumn.Activate
     With xRgUni_Account.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Account # column END
        
'Find Fup_count # column
    Dim xRg_Fup_count As Range
    Dim xRgUni_Fup_count As Range
    Dim xAddress_Fup_count As String
    Dim xStr_Fup_count As String
    On Error Resume Next
    xStr_Fup_count = "FollowUp Count"
    Set xRg_Fup_count = ActiveSheet.UsedRange.Find(xStr_Fup_count, , xlValues, xlWhole, , , True)
    If Not xRg_Fup_count Is Nothing Then
        xAddress_Fup_count = xRg_Fup_count.Address
        Do
            Set xRg_Fup_count = ActiveSheet.UsedRange.FindNext(xRg_Fup_count)
            If xRgUni_Fup_count Is Nothing Then
                Set xRgUni_Fup_count = xRg_Fup_count
            Else
                Set xRgUni_Fup_count = Application.Union(xRgUni_Fup_count, xRg_Fup_count)
            End If
        Loop While (Not xRg_Fup_count Is Nothing) And (xRg_Fup_count.Address <> xAddress_Fup_count)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Fup_count.EntireColumn.Activate
     With xRgUni_Fup_count.EntireColumn
        .ColumnWidth = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .FormatConditions.Delete
    End With
    
    'Color foormating
        'Add first rule
        xRgUni_Fup_count.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
                Formula1:="=2", Formula2:=" "
        xRgUni_Fup_count.EntireColumn.FormatConditions(1).Interior.Color = RGB(255, 0, 0)
        xRgUni_Fup_count.EntireColumn.FormatConditions(1).Font.Color = RGB(255, 255, 255)
        xRgUni_Fup_count.EntireColumn.FormatConditions(1).Font.Bold = True
                'Add second rule
        xRgUni_Fup_count.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
                Formula1:="=1"
        xRgUni_Fup_count.EntireColumn.FormatConditions(2).Interior.Color = RGB(255, 128, 0)
        'Add third rule
        'xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        '       Formula1:="=0"
        'xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = vbYellow

  ' end colors
  
  
'Find Fup_count # column  end

'Find Account_Type # column
    Dim xRg_Account_Type As Range
    Dim xRgUni_Account_Type As Range
    Dim xAddress_Account_Type As String
    Dim xStr_Account_Type As String
    On Error Resume Next
    xStr_Account_Type = "Account Type"
    Set xRg_Account_Type = ActiveSheet.UsedRange.Find(xStr_Account_Type, , xlValues, xlWhole, , , True)
    If Not xRg_Account_Type Is Nothing Then
        xAddress_Account_Type = xRg_Account_Type.Address
        Do
            Set xRg_Account_Type = ActiveSheet.UsedRange.FindNext(xRg_Account_Type)
            If xRgUni_Account_Type Is Nothing Then
                Set xRgUni_Account_Type = xRg_Account_Type
            Else
                Set xRgUni_Account_Type = Application.Union(xRgUni_Account_Type, xRg_Account_Type)
            End If
        Loop While (Not xRg_Account_Type Is Nothing) And (xRg_Account_Type.Address <> xAddress_Account_Type)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Account_Type.EntireColumn.Activate
     With xRgUni_Account_Type.EntireColumn
        .ColumnWidth = 9.43
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Account_Type # column  end

    
'Find age # column
    Dim xRg_age As Range
    Dim xRgUni_age As Range
    Dim xAddress_age As String
    Dim xStr_age As String
    On Error Resume Next
    xStr_age = "Age"
    Set xRg_age = ActiveSheet.UsedRange.Find(xStr_age, , xlValues, xlWhole, , , True)
    If Not xRg_age Is Nothing Then
        xAddress_age = xRg_age.Address
        Do
            Set xRg_age = ActiveSheet.UsedRange.FindNext(xRg_age)
            If xRgUni_age Is Nothing Then
                Set xRgUni_age = xRg_age
            Else
                Set xRgUni_age = Application.Union(xRgUni_age, xRg_age)
            End If
        Loop While (Not xRg_age Is Nothing) And (xRg_age.Address <> xAddress_age)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_age.EntireColumn.Activate
     With xRgUni_age.EntireColumn
        .ColumnWidth = 4.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .FormatConditions.Delete
    End With
'Color foormating
'Add first rule
		xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlBetween, _
				Formula1:="=3", Formula2:=" "
		xRgUni_age.EntireColumn.FormatConditions(1).Interior.Color = RGB(255,105,97)
		xRgUni_age.EntireColumn.FormatConditions(1).Font.Color = RGB(0, 0, 0)
        xRgUni_age.EntireColumn.FormatConditions(1).Font.Bold = False
		'Add second rule
		xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
				Formula1:="=2"
		xRgUni_age.EntireColumn.FormatConditions(2).Interior.Color = RGB(255, 128, 0)
		'Add third rule
		'xRgUni_age.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
			'	Formula1:="=0"
		'xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = vbWhite
		'Add fourth rule
		xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
				Formula1:="=1"
		xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = RGB(225, 229, 204)
		' Add fifth rule
		'xRgUni_age.EntireColumn.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
		'		Formula1:="=0"
		'xRgUni_age.EntireColumn.FormatConditions(3).Interior.Color = vbYellow

'End of SR age
    

'Find Reg_phone # column
    Dim xRg_Reg_phone As Range
    Dim xRgUni_Reg_phone As Range
    Dim xAddress_Reg_phone As String
    Dim xStr_Reg_phone As String
    On Error Resume Next
    xStr_Reg_phone = "Registered Phone"
    Set xRg_Reg_phone = ActiveSheet.UsedRange.Find(xStr_Reg_phone, , xlValues, xlWhole, , , True)
    If Not xRg_Reg_phone Is Nothing Then
        xAddress_Reg_phone = xRg_Reg_phone.Address
        Do
            Set xRg_Reg_phone = ActiveSheet.UsedRange.FindNext(xRg_Reg_phone)
            If xRgUni_Reg_phone Is Nothing Then
                Set xRgUni_Reg_phone = xRg_Reg_phone
            Else
                Set xRgUni_Reg_phone = Application.Union(xRgUni_Reg_phone, xRg_Reg_phone)
            End If
        Loop While (Not xRg_Reg_phone Is Nothing) And (xRg_Reg_phone.Address <> xAddress_Reg_phone)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Reg_phone.EntireColumn.Activate
     With xRgUni_Reg_phone.EntireColumn
        .ColumnWidth = 11.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Alt_phone # column
    Dim xRg_Alt_phone As Range
    Dim xRgUni_Alt_phone As Range
    Dim xAddress_Alt_phone As String
    Dim xStr_Alt_phone As String
    On Error Resume Next
    xStr_Alt_phone = "Alternate Phone"
    Set xRg_Alt_phone = ActiveSheet.UsedRange.Find(xStr_Alt_phone, , xlValues, xlWhole, , , True)
    If Not xRg_Alt_phone Is Nothing Then
        xAddress_Alt_phone = xRg_Alt_phone.Address
        Do
            Set xRg_Alt_phone = ActiveSheet.UsedRange.FindNext(xRg_Alt_phone)
            If xRgUni_Alt_phone Is Nothing Then
                Set xRgUni_Alt_phone = xRg_Alt_phone
            Else
                Set xRgUni_Alt_phone = Application.Union(xRgUni_Alt_phone, xRg_Alt_phone)
            End If
        Loop While (Not xRg_Alt_phone Is Nothing) And (xRg_Alt_phone.Address <> xAddress_Alt_phone)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Alt_phone.EntireColumn.Activate
     With xRgUni_Alt_phone.EntireColumn
        .ColumnWidth = 11.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

'Find Product_Category # column
    Dim xRg_Product_Category As Range
    Dim xRgUni_Product_Category As Range
    Dim xAddress_Product_Category As String
    Dim xStr_Product_Category As String
    On Error Resume Next
    xStr_Product_Category = "Product Category"
    Set xRg_Product_Category = ActiveSheet.UsedRange.Find(xStr_Product_Category, , xlValues, xlWhole, , , True)
    If Not xRg_Product_Category Is Nothing Then
        xAddress_Product_Category = xRg_Product_Category.Address
        Do
            Set xRg_Product_Category = ActiveSheet.UsedRange.FindNext(xRg_Product_Category)
            If xRgUni_Product_Category Is Nothing Then
                Set xRgUni_Product_Category = xRg_Product_Category
            Else
                Set xRgUni_Product_Category = Application.Union(xRgUni_Product_Category, xRg_Product_Category)
            End If
        Loop While (Not xRg_Product_Category Is Nothing) And (xRg_Product_Category.Address <> xAddress_Product_Category)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Product_Category.EntireColumn.Activate
     With xRgUni_Product_Category.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Unit_Status # column
    Dim xRg_Unit_Status As Range
    Dim xRgUni_Unit_Status As Range
    Dim xAddress_Unit_Status As String
    Dim xStr_Unit_Status As String
    On Error Resume Next
    xStr_Unit_Status = "Unit Status"
    Set xRg_Unit_Status = ActiveSheet.UsedRange.Find(xStr_Unit_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Unit_Status Is Nothing Then
        xAddress_Unit_Status = xRg_Unit_Status.Address
        Do
            Set xRg_Unit_Status = ActiveSheet.UsedRange.FindNext(xRg_Unit_Status)
            If xRgUni_Unit_Status Is Nothing Then
                Set xRgUni_Unit_Status = xRg_Unit_Status
            Else
                Set xRgUni_Unit_Status = Application.Union(xRgUni_Unit_Status, xRg_Unit_Status)
            End If
        Loop While (Not xRg_Unit_Status Is Nothing) And (xRg_Unit_Status.Address <> xAddress_Unit_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Unit_Status.EntireColumn.Activate
     With xRgUni_Unit_Status.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Status # column
    Dim xRg_Status As Range
    Dim xRgUni_Status As Range
    Dim xAddress_Status As String
    Dim xStr_Status As String
    On Error Resume Next
    xStr_Status = "Status"
    Set xRg_Status = ActiveSheet.UsedRange.Find(xStr_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Status Is Nothing Then
        xAddress_Status = xRg_Status.Address
        Do
            Set xRg_Status = ActiveSheet.UsedRange.FindNext(xRg_Status)
            If xRgUni_Status Is Nothing Then
                Set xRgUni_Status = xRg_Status
            Else
                Set xRgUni_Status = Application.Union(xRgUni_Status, xRg_Status)
            End If
        Loop While (Not xRg_Status Is Nothing) And (xRg_Status.Address <> xAddress_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Status.EntireColumn.Activate
     With xRgUni_Status.EntireColumn
        .ColumnWidth = 7.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

        
'Find Sub_Status # column
    Dim xRg_Sub_Status As Range
    Dim xRgUni_Sub_Status As Range
    Dim xAddress_Sub_Status As String
    Dim xStr_Sub_Status As String
    On Error Resume Next
    xStr_Sub_Status = "Sub Status"
    Set xRg_Sub_Status = ActiveSheet.UsedRange.Find(xStr_Sub_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Sub_Status Is Nothing Then
        xAddress_Sub_Status = xRg_Sub_Status.Address
        Do
            Set xRg_Sub_Status = ActiveSheet.UsedRange.FindNext(xRg_Sub_Status)
            If xRgUni_Sub_Status Is Nothing Then
                Set xRgUni_Sub_Status = xRg_Sub_Status
            Else
                Set xRgUni_Sub_Status = Application.Union(xRgUni_Sub_Status, xRg_Sub_Status)
            End If
        Loop While (Not xRg_Sub_Status Is Nothing) And (xRg_Sub_Status.Address <> xAddress_Sub_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Sub_Status.EntireColumn.Activate
     With xRgUni_Sub_Status.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .FormatConditions.Delete
    End With
    
'Color foormating with string
'Re-opened calls
	With xRgUni_Sub_Status.EntireColumn.FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Re-Opened")
		With .Font
			.Bold = True
			.ColorIndex = 3
		End With
	End With

	With xRgUni_Sub_Status.EntireColumn.FormatConditions.Add(xlTextString, TextOperator:=xlContains, String:="Cancel Request Rejected")
		With .Font
			.Bold = True
			.ColorIndex = 3
		End With
	End With
        
'Find Key_Account_S # column
    Dim xRg_Key_Account_S As Range
    Dim xRgUni_Key_Account_S As Range
    Dim xAddress_Key_Account_S As String
    Dim xStr_Key_Account_S As String
    On Error Resume Next
    xStr_Key_Account_S = "Key Account (S)"
    Set xRg_Key_Account_S = ActiveSheet.UsedRange.Find(xStr_Key_Account_S, , xlValues, xlWhole, , , True)
    If Not xRg_Key_Account_S Is Nothing Then
        xAddress_Key_Account_S = xRg_Key_Account_S.Address
        Do
            Set xRg_Key_Account_S = ActiveSheet.UsedRange.FindNext(xRg_Key_Account_S)
            If xRgUni_Key_Account_S Is Nothing Then
                Set xRgUni_Key_Account_S = xRg_Key_Account_S
            Else
                Set xRgUni_Key_Account_S = Application.Union(xRgUni_Key_Account_S, xRg_Key_Account_S)
            End If
        Loop While (Not xRg_Key_Account_S Is Nothing) And (xRg_Key_Account_S.Address <> xAddress_Key_Account_S)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Key_Account_S.EntireColumn.Activate
     With xRgUni_Key_Account_S.EntireColumn
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

'Find Key_Account_P # column
    Dim xRg_Key_Account_P As Range
    Dim xRgUni_Key_Account_P As Range
    Dim xAddress_Key_Account_P As String
    Dim xStr_Key_Account_P As String
    On Error Resume Next
    xStr_Key_Account_P = "Key Account (P)"
    Set xRg_Key_Account_P = ActiveSheet.UsedRange.Find(xStr_Key_Account_P, , xlValues, xlWhole, , , True)
    If Not xRg_Key_Account_P Is Nothing Then
        xAddress_Key_Account_P = xRg_Key_Account_P.Address
        Do
            Set xRg_Key_Account_P = ActiveSheet.UsedRange.FindNext(xRg_Key_Account_P)
            If xRgUni_Key_Account_P Is Nothing Then
                Set xRgUni_Key_Account_P = xRg_Key_Account_P
            Else
                Set xRgUni_Key_Account_P = Application.Union(xRgUni_Key_Account_P, xRg_Key_Account_P)
            End If
        Loop While (Not xRg_Key_Account_P Is Nothing) And (xRg_Key_Account_P.Address <> xAddress_Key_Account_P)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Key_Account_P.EntireColumn.Activate
     With xRgUni_Key_Account_P.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Technician # column SF
    Dim xRg_Technician As Range
    Dim xRgUni_Technician As Range
    Dim xAddress_Technician As String
    Dim xStr_Technician As String
    On Error Resume Next
    xStr_Technician = "Owner"
    Set xRg_Technician = ActiveSheet.UsedRange.Find(xStr_Technician, , xlValues, xlWhole, , , True)
    If Not xRg_Technician Is Nothing Then
        xAddress_Technician = xRg_Technician.Address
        Do
            Set xRg_Technician = ActiveSheet.UsedRange.FindNext(xRg_Technician)
            If xRgUni_Technician Is Nothing Then
                Set xRgUni_Technician = xRg_Technician
            Else
                Set xRgUni_Technician = Application.Union(xRgUni_Technician, xRg_Technician)
            End If
        Loop While (Not xRg_Technician Is Nothing) And (xRg_Technician.Address <> xAddress_Technician)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Technician.EntireColumn.Activate
     With xRgUni_Technician.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Technician_ASM
    Dim xRg_Technician_ASM As Range
    Dim xRgUni_Technician_ASM As Range
    Dim xAddress_Technician_ASM As String
    Dim xStr_Technician_ASM As String
    On Error Resume Next
    xStr_Technician_ASM = "Technician"
    Set xRg_Technician_ASM = ActiveSheet.UsedRange.Find(xStr_Technician_ASM, , xlValues, xlWhole, , , True)
    If Not xRg_Technician_ASM Is Nothing Then
        xAddress_Technician_ASM = xRg_Technician_ASM.Address
        Do
            Set xRg_Technician_ASM = ActiveSheet.UsedRange.FindNext(xRg_Technician_ASM)
            If xRgUni_Technician_ASM Is Nothing Then
                Set xRgUni_Technician_ASM = xRg_Technician_ASM
            Else
                Set xRgUni_Technician_ASM = Application.Union(xRgUni_Technician_ASM, xRg_Technician_ASM)
            End If
        Loop While (Not xRg_Technician_ASM Is Nothing) And (xRg_Technician_ASM.Address <> xAddress_Technician_ASM)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Technician_ASM.EntireColumn.Activate
     With xRgUni_Technician_ASM.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find ASM
    Dim xRg_ASM As Range
    Dim xRgUni_ASM As Range
    Dim xAddress_ASM As String
    Dim xStr_ASM As String
    On Error Resume Next
    xStr_ASM = "ASM"
    Set xRg_ASM = ActiveSheet.UsedRange.Find(xStr_ASM, , xlValues, xlWhole, , , True)
    If Not xRg_ASM Is Nothing Then
        xAddress_ASM = xRg_ASM.Address
        Do
            Set xRg_ASM = ActiveSheet.UsedRange.FindNext(xRg_ASM)
            If xRgUni_ASM Is Nothing Then
                Set xRgUni_ASM = xRg_ASM
            Else
                Set xRgUni_ASM = Application.Union(xRgUni_ASM, xRg_ASM)
            End If
        Loop While (Not xRg_ASM Is Nothing) And (xRg_ASM.Address <> xAddress_ASM)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_ASM.EntireColumn.Activate
     With xRgUni_ASM.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
        
'Find Service_Agent
    Dim xRg_Service_Agent As Range
    Dim xRgUni_Service_Agent As Range
    Dim xAddress_Service_Agent As String
    Dim xStr_Service_Agent As String
    On Error Resume Next
    xStr_Service_Agent = "Service Agent"
    Set xRg_Service_Agent = ActiveSheet.UsedRange.Find(xStr_Service_Agent, , xlValues, xlWhole, , , True)
    If Not xRg_Service_Agent Is Nothing Then
        xAddress_Service_Agent = xRg_Service_Agent.Address
        Do
            Set xRg_Service_Agent = ActiveSheet.UsedRange.FindNext(xRg_Service_Agent)
            If xRgUni_Service_Agent Is Nothing Then
                Set xRgUni_Service_Agent = xRg_Service_Agent
            Else
                Set xRgUni_Service_Agent = Application.Union(xRgUni_Service_Agent, xRg_Service_Agent)
            End If
        Loop While (Not xRg_Service_Agent Is Nothing) And (xRg_Service_Agent.Address <> xAddress_Service_Agent)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Service_Agent.EntireColumn.Activate
     With xRgUni_Service_Agent.EntireColumn
        .ColumnWidth = 25
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
        
'Find Description
    Dim xRg_Description As Range
    Dim xRgUni_Description As Range
    Dim xAddress_Description As String
    Dim xStr_Description As String
    On Error Resume Next
    xStr_Description = "Description"
    Set xRg_Description = ActiveSheet.UsedRange.Find(xStr_Description, , xlValues, xlWhole, , , True)
    If Not xRg_Description Is Nothing Then
        xAddress_Description = xRg_Description.Address
        Do
            Set xRg_Description = ActiveSheet.UsedRange.FindNext(xRg_Description)
            If xRgUni_Description Is Nothing Then
                Set xRgUni_Description = xRg_Description
            Else
                Set xRgUni_Description = Application.Union(xRgUni_Description, xRg_Description)
            End If
        Loop While (Not xRg_Description Is Nothing) And (xRg_Description.Address <> xAddress_Description)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Description.EntireColumn.Activate
     With xRgUni_Description.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

'Find Audit_Status
    Dim xRg_Audit_Status As Range
    Dim xRgUni_Audit_Status As Range
    Dim xAddress_Audit_Status As String
    Dim xStr_Audit_Status As String
    On Error Resume Next
    xStr_Audit_Status = "Audit Status"
    Set xRg_Audit_Status = ActiveSheet.UsedRange.Find(xStr_Audit_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Audit_Status Is Nothing Then
        xAddress_Audit_Status = xRg_Audit_Status.Address
        Do
            Set xRg_Audit_Status = ActiveSheet.UsedRange.FindNext(xRg_Audit_Status)
            If xRgUni_Audit_Status Is Nothing Then
                Set xRgUni_Audit_Status = xRg_Audit_Status
            Else
                Set xRgUni_Audit_Status = Application.Union(xRgUni_Audit_Status, xRg_Audit_Status)
            End If
        Loop While (Not xRg_Audit_Status Is Nothing) And (xRg_Audit_Status.Address <> xAddress_Audit_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Audit_Status.EntireColumn.Activate
     With xRgUni_Audit_Status.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

'Find Brand_Identifier
    Dim xRg_Brand_Identifier As Range
    Dim xRgUni_Brand_Identifier As Range
    Dim xAddress_Brand_Identifier As String
    Dim xStr_Brand_Identifier As String
    On Error Resume Next
    xStr_Brand_Identifier = "Brand Identifier"
    Set xRg_Brand_Identifier = ActiveSheet.UsedRange.Find(xStr_Brand_Identifier, , xlValues, xlWhole, , , True)
    If Not xRg_Brand_Identifier Is Nothing Then
        xAddress_Brand_Identifier = xRg_Brand_Identifier.Address
        Do
            Set xRg_Brand_Identifier = ActiveSheet.UsedRange.FindNext(xRg_Brand_Identifier)
            If xRgUni_Brand_Identifier Is Nothing Then
                Set xRgUni_Brand_Identifier = xRg_Brand_Identifier
            Else
                Set xRgUni_Brand_Identifier = Application.Union(xRgUni_Brand_Identifier, xRg_Brand_Identifier)
            End If
        Loop While (Not xRg_Brand_Identifier Is Nothing) And (xRg_Brand_Identifier.Address <> xAddress_Brand_Identifier)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Brand_Identifier.EntireColumn.Activate
     With xRgUni_Brand_Identifier.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Penalty_Order
    Dim xRg_Penalty_Order As Range
    Dim xRgUni_Penalty_Order As Range
    Dim xAddress_Penalty_Order As String
    Dim xStr_Penalty_Order As String
    On Error Resume Next
    xStr_Penalty_Order = "Penalty Order#"
    Set xRg_Penalty_Order = ActiveSheet.UsedRange.Find(xStr_Penalty_Order, , xlValues, xlWhole, , , True)
    If Not xRg_Penalty_Order Is Nothing Then
        xAddress_Penalty_Order = xRg_Penalty_Order.Address
        Do
            Set xRg_Penalty_Order = ActiveSheet.UsedRange.FindNext(xRg_Penalty_Order)
            If xRgUni_Penalty_Order Is Nothing Then
                Set xRgUni_Penalty_Order = xRg_Penalty_Order
            Else
                Set xRgUni_Penalty_Order = Application.Union(xRgUni_Penalty_Order, xRg_Penalty_Order)
            End If
        Loop While (Not xRg_Penalty_Order Is Nothing) And (xRg_Penalty_Order.Address <> xAddress_Penalty_Order)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Penalty_Order.EntireColumn.Activate
     With xRgUni_Penalty_Order.EntireColumn
        .ColumnWidth = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Pincode
    Dim xRg_Pincode As Range
    Dim xRgUni_Pincode As Range
    Dim xAddress_Pincode As String
    Dim xStr_Pincode As String
    On Error Resume Next
    xStr_Pincode = "Pincode"
    Set xRg_Pincode = ActiveSheet.UsedRange.Find(xStr_Pincode, , xlValues, xlWhole, , , True)
    If Not xRg_Pincode Is Nothing Then
        xAddress_Pincode = xRg_Pincode.Address
        Do
            Set xRg_Pincode = ActiveSheet.UsedRange.FindNext(xRg_Pincode)
            If xRgUni_Pincode Is Nothing Then
                Set xRgUni_Pincode = xRg_Pincode
            Else
                Set xRgUni_Pincode = Application.Union(xRgUni_Pincode, xRg_Pincode)
            End If
        Loop While (Not xRg_Pincode Is Nothing) And (xRg_Pincode.Address <> xAddress_Pincode)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Pincode.EntireColumn.Activate
     With xRgUni_Pincode.EntireColumn
        .ColumnWidth = 7.5
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Area
    Dim xRg_Area As Range
    Dim xRgUni_Area As Range
    Dim xAddress_Area As String
    Dim xStr_Area As String
    On Error Resume Next
    xStr_Area = "Area"
    Set xRg_Area = ActiveSheet.UsedRange.Find(xStr_Area, , xlValues, xlWhole, , , True)
    If Not xRg_Area Is Nothing Then
        xAddress_Area = xRg_Area.Address
        Do
            Set xRg_Area = ActiveSheet.UsedRange.FindNext(xRg_Area)
            If xRgUni_Area Is Nothing Then
                Set xRgUni_Area = xRg_Area
            Else
                Set xRgUni_Area = Application.Union(xRgUni_Area, xRg_Area)
            End If
        Loop While (Not xRg_Area Is Nothing) And (xRg_Area.Address <> xAddress_Area)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Area.EntireColumn.Activate
     With xRgUni_Area.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find City
    Dim xRg_City As Range
    Dim xRgUni_City As Range
    Dim xAddress_City As String
    Dim xStr_City As String
    On Error Resume Next
    xStr_City = "City"
    Set xRg_City = ActiveSheet.UsedRange.Find(xStr_City, , xlValues, xlWhole, , , True)
    If Not xRg_City Is Nothing Then
        xAddress_City = xRg_City.Address
        Do
            Set xRg_City = ActiveSheet.UsedRange.FindNext(xRg_City)
            If xRgUni_City Is Nothing Then
                Set xRgUni_City = xRg_City
            Else
                Set xRgUni_City = Application.Union(xRgUni_City, xRg_City)
            End If
        Loop While (Not xRg_City Is Nothing) And (xRg_City.Address <> xAddress_City)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_City.EntireColumn.Activate
     With xRgUni_City.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find District
    Dim xRg_District As Range
    Dim xRgUni_District As Range
    Dim xAddress_District As String
    Dim xStr_District As String
    On Error Resume Next
    xStr_District = "District"
    Set xRg_District = ActiveSheet.UsedRange.Find(xStr_District, , xlValues, xlWhole, , , True)
    If Not xRg_District Is Nothing Then
        xAddress_District = xRg_District.Address
        Do
            Set xRg_District = ActiveSheet.UsedRange.FindNext(xRg_District)
            If xRgUni_District Is Nothing Then
                Set xRgUni_District = xRg_District
            Else
                Set xRgUni_District = Application.Union(xRgUni_District, xRg_District)
            End If
        Loop While (Not xRg_District Is Nothing) And (xRg_District.Address <> xAddress_District)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_District.EntireColumn.Activate
     With xRgUni_District.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Processing_Status
    Dim xRg_Processing_Status As Range
    Dim xRgUni_Processing_Status As Range
    Dim xAddress_Processing_Status As String
    Dim xStr_Processing_Status As String
    On Error Resume Next
    xStr_Processing_Status = "Processing Status"
    Set xRg_Processing_Status = ActiveSheet.UsedRange.Find(xStr_Processing_Status, , xlValues, xlWhole, , , True)
    If Not xRg_Processing_Status Is Nothing Then
        xAddress_Processing_Status = xRg_Processing_Status.Address
        Do
            Set xRg_Processing_Status = ActiveSheet.UsedRange.FindNext(xRg_Processing_Status)
            If xRgUni_Processing_Status Is Nothing Then
                Set xRgUni_Processing_Status = xRg_Processing_Status
            Else
                Set xRgUni_Processing_Status = Application.Union(xRgUni_Processing_Status, xRg_Processing_Status)
            End If
        Loop While (Not xRg_Processing_Status Is Nothing) And (xRg_Processing_Status.Address <> xAddress_Processing_Status)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Processing_Status.EntireColumn.Activate
     With xRgUni_Processing_Status.EntireColumn
        .ColumnWidth = 15
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Processed
    Dim xRg_Processed As Range
    Dim xRgUni_Processed As Range
    Dim xAddress_Processed As String
    Dim xStr_Processed As String
    On Error Resume Next
    xStr_Processed = "Processed"
    Set xRg_Processed = ActiveSheet.UsedRange.Find(xStr_Processed, , xlValues, xlWhole, , , True)
    If Not xRg_Processed Is Nothing Then
        xAddress_Processed = xRg_Processed.Address
        Do
            Set xRg_Processed = ActiveSheet.UsedRange.FindNext(xRg_Processed)
            If xRgUni_Processed Is Nothing Then
                Set xRgUni_Processed = xRg_Processed
            Else
                Set xRgUni_Processed = Application.Union(xRgUni_Processed, xRg_Processed)
            End If
        Loop While (Not xRg_Processed Is Nothing) And (xRg_Processed.Address <> xAddress_Processed)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Processed.EntireColumn.Activate
     With xRgUni_Processed.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
        
'Find Expense_Count
    Dim xRg_Expense_Count As Range
    Dim xRgUni_Expense_Count As Range
    Dim xAddress_Expense_Count As String
    Dim xStr_Expense_Count As String
    On Error Resume Next
    xStr_Expense_Count = "Expense Count"
    Set xRg_Expense_Count = ActiveSheet.UsedRange.Find(xStr_Expense_Count, , xlValues, xlWhole, , , True)
    If Not xRg_Expense_Count Is Nothing Then
        xAddress_Expense_Count = xRg_Expense_Count.Address
        Do
            Set xRg_Expense_Count = ActiveSheet.UsedRange.FindNext(xRg_Expense_Count)
            If xRgUni_Expense_Count Is Nothing Then
                Set xRgUni_Expense_Count = xRg_Expense_Count
            Else
                Set xRgUni_Expense_Count = Application.Union(xRgUni_Expense_Count, xRg_Expense_Count)
            End If
        Loop While (Not xRg_Expense_Count Is Nothing) And (xRg_Expense_Count.Address <> xAddress_Expense_Count)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Expense_Count.EntireColumn.Activate
     With xRgUni_Expense_Count.EntireColumn
        .ColumnWidth = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Closed_Date
    Dim xRg_Closed_Date As Range
    Dim xRgUni_Closed_Date As Range
    Dim xAddress_Closed_Date As String
    Dim xStr_Closed_Date As String
    On Error Resume Next
    xStr_Closed_Date = "Closed Date"
    Set xRg_Closed_Date = ActiveSheet.UsedRange.Find(xStr_Closed_Date, , xlValues, xlWhole, , , True)
    If Not xRg_Closed_Date Is Nothing Then
        xAddress_Closed_Date = xRg_Closed_Date.Address
        Do
            Set xRg_Closed_Date = ActiveSheet.UsedRange.FindNext(xRg_Closed_Date)
            If xRgUni_Closed_Date Is Nothing Then
                Set xRgUni_Closed_Date = xRg_Closed_Date
            Else
                Set xRgUni_Closed_Date = Application.Union(xRgUni_Closed_Date, xRg_Closed_Date)
            End If
        Loop While (Not xRg_Closed_Date Is Nothing) And (xRg_Closed_Date.Address <> xAddress_Closed_Date)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Closed_Date.EntireColumn.Activate
     With xRgUni_Closed_Date.EntireColumn
        .ColumnWidth = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
        .NumberFormat = "dd-mm-yyyy"
    End With
        
'Find Serial_Split
    Dim xRg_Serial_Split As Range
    Dim xRgUni_Serial_Split As Range
    Dim xAddress_Serial_Split As String
    Dim xStr_Serial_Split As String
    On Error Resume Next
    xStr_Serial_Split = "Serial #(Split)"
    Set xRg_Serial_Split = ActiveSheet.UsedRange.Find(xStr_Serial_Split, , xlValues, xlWhole, , , True)
    If Not xRg_Serial_Split Is Nothing Then
        xAddress_Serial_Split = xRg_Serial_Split.Address
        Do
            Set xRg_Serial_Split = ActiveSheet.UsedRange.FindNext(xRg_Serial_Split)
            If xRgUni_Serial_Split Is Nothing Then
                Set xRgUni_Serial_Split = xRg_Serial_Split
            Else
                Set xRgUni_Serial_Split = Application.Union(xRgUni_Serial_Split, xRg_Serial_Split)
            End If
        Loop While (Not xRg_Serial_Split Is Nothing) And (xRg_Serial_Split.Address <> xAddress_Serial_Split)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Serial_Split.EntireColumn.Activate
     With xRgUni_Serial_Split.EntireColumn
        .ColumnWidth = 17
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
        
'Find Serial
    Dim xRg_Serial As Range
    Dim xRgUni_Serial As Range
    Dim xAddress_Serial As String
    Dim xStr_Serial As String
    On Error Resume Next
    xStr_Serial = "Serial #"
    Set xRg_Serial = ActiveSheet.UsedRange.Find(xStr_Serial, , xlValues, xlWhole, , , True)
    If Not xRg_Serial Is Nothing Then
        xAddress_Serial = xRg_Serial.Address
        Do
            Set xRg_Serial = ActiveSheet.UsedRange.FindNext(xRg_Serial)
            If xRgUni_Serial Is Nothing Then
                Set xRgUni_Serial = xRg_Serial
            Else
                Set xRgUni_Serial = Application.Union(xRgUni_Serial, xRg_Serial)
            End If
        Loop While (Not xRg_Serial Is Nothing) And (xRg_Serial.Address <> xAddress_Serial)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Serial.EntireColumn.Activate
     With xRgUni_Serial.EntireColumn
        .ColumnWidth = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
'Find Model
    Dim xRg_Model As Range
    Dim xRgUni_Model As Range
    Dim xAddress_Model As String
    Dim xStr_Model As String
    On Error Resume Next
    xStr_Model = "Model"
    Set xRg_Model = ActiveSheet.UsedRange.Find(xStr_Model, , xlValues, xlWhole, , , True)
    If Not xRg_Model Is Nothing Then
        xAddress_Model = xRg_Model.Address
        Do
            Set xRg_Model = ActiveSheet.UsedRange.FindNext(xRg_Model)
            If xRgUni_Model Is Nothing Then
                Set xRgUni_Model = xRg_Model
            Else
                Set xRgUni_Model = Application.Union(xRgUni_Model, xRg_Model)
            End If
        Loop While (Not xRg_Model Is Nothing) And (xRg_Model.Address <> xAddress_Model)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Model.EntireColumn.Activate
     With xRgUni_Model.EntireColumn
        .ColumnWidth = 20
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
        
'Find Assign_Date
    Dim xRg_Assign_Date As Range
    Dim xRgUni_Assign_Date As Range
    Dim xAddress_Assign_Date As String
    Dim xStr_Assign_Date As String
    On Error Resume Next
    xStr_Assign_Date = "Assign Date"
    Set xRg_Assign_Date = ActiveSheet.UsedRange.Find(xStr_Assign_Date, , xlValues, xlWhole, , , True)
    If Not xRg_Assign_Date Is Nothing Then
        xAddress_Assign_Date = xRg_Assign_Date.Address
        Do
            Set xRg_Assign_Date = ActiveSheet.UsedRange.FindNext(xRg_Assign_Date)
            If xRgUni_Assign_Date Is Nothing Then
                Set xRgUni_Assign_Date = xRg_Assign_Date
            Else
                Set xRgUni_Assign_Date = Application.Union(xRgUni_Assign_Date, xRg_Assign_Date)
            End If
        Loop While (Not xRg_Assign_Date Is Nothing) And (xRg_Assign_Date.Address <> xAddress_Assign_Date)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Assign_Date.EntireColumn.Activate
     With xRgUni_Assign_Date.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .NumberFormat = "dd-mm-yyyy"
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
        
        
'Find Open_Date
    Dim xRg_Open_Date As Range
    Dim xRgUni_Open_Date As Range
    Dim xAddress_Open_Date As String
    Dim xStr_Open_Date As String
    On Error Resume Next
    xStr_Open_Date = "Open Date"
    Set xRg_Open_Date = ActiveSheet.UsedRange.Find(xStr_Open_Date, , xlValues, xlWhole, , , True)
    If Not xRg_Open_Date Is Nothing Then
        xAddress_Open_Date = xRg_Open_Date.Address
        Do
            Set xRg_Open_Date = ActiveSheet.UsedRange.FindNext(xRg_Open_Date)
            If xRgUni_Open_Date Is Nothing Then
                Set xRgUni_Open_Date = xRg_Open_Date
            Else
                Set xRgUni_Open_Date = Application.Union(xRgUni_Open_Date, xRg_Open_Date)
            End If
        Loop While (Not xRg_Open_Date Is Nothing) And (xRg_Open_Date.Address <> xAddress_Open_Date)
    End If
    'Application.CutCopyMode = False ' don't want an existing operation to interfere
    xRgUni_Open_Date.EntireColumn.Activate
     With xRgUni_Open_Date.EntireColumn
        .ColumnWidth = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .NumberFormat = "dd-mm-yyyy"
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With

'Find Action_Taken # column
    Dim xRg_Action_Taken As Range
    Dim xRgUni_Action_Taken As Range
    Dim xAddress_Action_Taken As String
    Dim xStr_Action_Taken As String
    On Error Resume Next
    xStr_Action_Taken = "Action Taken"
    Set xRg_Action_Taken = ActiveSheet.UsedRange.Find(xStr_Action_Taken, , xlValues, xlWhole, , , True)
    If Not xRg_Action_Taken Is Nothing Then
        xAddress_Action_Taken = xRg_Action_Taken.Address
        Do
            Set xRg_Action_Taken = ActiveSheet.UsedRange.FindNext(xRg_Action_Taken)
            If xRgUni_Action_Taken Is Nothing Then
                Set xRgUni_Action_Taken = xRg_Action_Taken
            Else
                Set xRgUni_Action_Taken = Application.Union(xRgUni_Action_Taken, xRg_Action_Taken)
            End If
        Loop While (Not xRg_Action_Taken Is Nothing) And (xRg_Action_Taken.Address <> xAddress_Action_Taken)
    End If
    xRgUni_Action_Taken.EntireColumn.Activate
     With xRgUni_Action_Taken.EntireColumn
        .AutoFit
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Action_Taken # column End

'Find Gas_Charge_Done_Flag # column
    Dim xRg_Gas_Charge_Done_Flag As Range
    Dim xRgUni_Gas_Charge_Done_Flag As Range
    Dim xAddress_Gas_Charge_Done_Flag As String
    Dim xStr_Gas_Charge_Done_Flag As String
    On Error Resume Next
    xStr_Gas_Charge_Done_Flag = "Gas Charge Done Flag"
    Set xRg_Gas_Charge_Done_Flag = ActiveSheet.UsedRange.Find(xStr_Gas_Charge_Done_Flag, , xlValues, xlWhole, , , True)
    If Not xRg_Gas_Charge_Done_Flag Is Nothing Then
        xAddress_Gas_Charge_Done_Flag = xRg_Gas_Charge_Done_Flag.Address
        Do
            Set xRg_Gas_Charge_Done_Flag = ActiveSheet.UsedRange.FindNext(xRg_Gas_Charge_Done_Flag)
            If xRgUni_Gas_Charge_Done_Flag Is Nothing Then
                Set xRgUni_Gas_Charge_Done_Flag = xRg_Gas_Charge_Done_Flag
            Else
                Set xRgUni_Gas_Charge_Done_Flag = Application.Union(xRgUni_Gas_Charge_Done_Flag, xRg_Gas_Charge_Done_Flag)
            End If
        Loop While (Not xRg_Gas_Charge_Done_Flag Is Nothing) And (xRg_Gas_Charge_Done_Flag.Address <> xAddress_Gas_Charge_Done_Flag)
    End If
    xRgUni_Gas_Charge_Done_Flag.EntireColumn.Activate
     With xRgUni_Gas_Charge_Done_Flag.EntireColumn
        .ColumnWidth = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Gas_Charge_Done_Flag # column End

'Find Part_Replaced_Flag # column
    Dim xRg_Part_Replaced_Flag As Range
    Dim xRgUni_Part_Replaced_Flag As Range
    Dim xAddress_Part_Replaced_Flag As String
    Dim xStr_Part_Replaced_Flag As String
    On Error Resume Next
    xStr_Part_Replaced_Flag = "Part Replaced Flag"
    Set xRg_Part_Replaced_Flag = ActiveSheet.UsedRange.Find(xStr_Part_Replaced_Flag, , xlValues, xlWhole, , , True)
    If Not xRg_Part_Replaced_Flag Is Nothing Then
        xAddress_Part_Replaced_Flag = xRg_Part_Replaced_Flag.Address
        Do
            Set xRg_Part_Replaced_Flag = ActiveSheet.UsedRange.FindNext(xRg_Part_Replaced_Flag)
            If xRgUni_Part_Replaced_Flag Is Nothing Then
                Set xRgUni_Part_Replaced_Flag = xRg_Part_Replaced_Flag
            Else
                Set xRgUni_Part_Replaced_Flag = Application.Union(xRgUni_Part_Replaced_Flag, xRg_Part_Replaced_Flag)
            End If
        Loop While (Not xRg_Part_Replaced_Flag Is Nothing) And (xRg_Part_Replaced_Flag.Address <> xAddress_Part_Replaced_Flag)
    End If
    xRgUni_Part_Replaced_Flag.EntireColumn.Activate
     With xRgUni_Part_Replaced_Flag.EntireColumn
        .ColumnWidth = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlCenter
        .MergeCells = False
    End With
'Find Part_Replaced_Flag # column End

'Add Borders
 With ActiveSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
 End With
  
'Color For Top row
With ActiveSheet.Range("A1", Cells(1, Columns.Count).End(xlToRight)).SpecialCells(xlCellTypeConstants)
  .Interior.ColorIndex = 6
  .Font.Bold = True
End With
  xRgUni_SR.Activate
  
  ' Page layout set up
  With ActiveSheet.PageSetup
     .Orientation = xlLandscape
     .PaperSize = xlPaperA4
     '.Zoom = 80
     .Zoom = False
     .FitToPagesTall = False
     .FitToPagesWide = False
     .LeftMargin = Application.InchesToPoints(0.35)
    .RightMargin = Application.InchesToPoints(0.35)
    .TopMargin = Application.InchesToPoints(0.35)
    .BottomMargin = Application.InchesToPoints(0.35)
     .HeaderMargin = Application.InchesToPoints(0.35)
     .FooterMargin = Application.InchesToPoints(0.35)
     
End With
End Sub
