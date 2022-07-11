Attribute VB_Name = "FormControl"
Private List() As Control
Private curr_obj As Object
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double

Private Type Control
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
End Type

Public Sub ResizeForm(frm As Form)
    'Set the forms height
    'frm.height = 10000
    'Set the forms width
    'frm.width = Screen.width / 1.2
    'Resize all of the controls
    'based on the forms new size
    ResizeControls frm
End Sub

Public Sub GetLocation(frm As Form)
Dim i As Integer
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function.
 
'Loop through each control
For Each curr_obj In frm
'Resize the Array by 1, and preserve
'the original objects in the array
    'If curr_obj = FileMenuBar Then
        On Error GoTo ErrorH
        ReDim Preserve List(i)
        With List(i)
                .Name = curr_obj.Name
                .Left = curr_obj.Left
                .Top = curr_obj.Top
                .width = curr_obj.width
                .height = curr_obj.height
        End With
Skipper:
        i = i + 1
ErrorH:
    If Err.Number = 438 Then
' do nothing
    Err.Clear
    End If
    Resume Next
    'End If
Next curr_obj
 
'   This is what the object sizes will be compared to on rescaling.
        iHeight = frm.ScaleHeight
        iWidth = frm.ScaleWidth
End Sub

Public Function SetFontSize() As Integer
    'Make sure x_size is greater than 0
    If Int(x_size) > 0 Then
    'Set the font size
        SetFontSize = Int(x_size * 8)
    End If
End Function

Public Sub ResizeControls(frm As Form)
    On Error Resume Next
    Dim i As Integer
    '   Get ratio of initial form size to current form size
    x_size = frm.height / iHeight
    y_size = frm.width / iWidth
    
    'Loop though all the control objects in the array
    'based on the upper bound of the # of controls
    For i = 0 To UBound(List)
        frm(List(i).Name).Left = List(i).Left * y_size
        frm(List(i).Name).width = List(i).width * y_size
        frm(List(i).Name).height = List(i).height * x_size
        frm(List(i).Name).Top = List(i).Top * x_size
    Next
    Err.Clear: On Error GoTo 0
End Sub

