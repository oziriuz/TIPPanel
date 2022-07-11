Attribute VB_Name = "Module1"
Option Explicit
Private List() As Controles
Private iHeight As Integer
Private iWidth As Integer
Private x_size As Double
Private y_size As Double


'*****************************************************************************************
'                           LICENSE INFORMATION
'*****************************************************************************************
'   FormControl Version 2.0
'   Code module for resizing a form based on screen size, then resizing the
'   controls based on the forms size
'
'   Copyright (C) 2007
'   Richard L. McCutchen
'   Email: richard@psychocoder.net
'   Created: AUG99
'
'   This program is free software: you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation, either version 3 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License
'   along with this program.  If not, see <http://www.gnu.org/licenses/>.
'*****************************************************************************************
Public Const RESIZEDEFAULT As String = "RESIZUPELEFTWIDTH"
Public Const RESIZEALL As String = "RESIZEALL"
Public Const NORESIZE As String = "NORESIZE"
Public Const RESIZEUP As String = "RESIZEUP"
Public Const RESIZELEFT As String = "RESIZELEFT"
Public Const RESIZEUPLEFT As String = "RESIZEUPLEFT"
Public Const RESIZEUPWIDTH As String = "RESIZUPWIDTH"
Public Const RESIZELEFTWIDTH As String = "RESIZUPLEFTWIDTH"
Public Const RESIZEUPLEFTWIDTH As String = "RESIZUPELEFTWIDTH"

Private Const Line = "Line"
Private Const Menu = "Menu"
Private Const SSTab = "SSTab"
Private Const Frame = "Frame"
Private Const Label = "Label"
Private Const Timer = "Timer"
Private Const TextBox = "TextBox"
Private Const ComboBox = "ComboBox"
Private Const CheckBox = "CheckBox"
Private Const TreeView = "TreeView"
Private Const ListView = "ListView"
Private Const MaskEdBox = "MaskEdBox"
Private Const ImageList = "ImageList"
Private Const RichTextBox = "RichTextBox"
Private Const CommandButton = "CommandButton"

Private Type Controles
    Index As Integer
    Name As String
    Left As Integer
    Top As Integer
    width As Integer
    height As Integer
End Type

Public Sub ResizeControls(frm As Form)
  Dim i, b As Integer
  Dim strTag As String
  Dim curr_obj As Control

  On Error Resume Next
'   Get ratio of initial form size to current form size
  x_size = frm.height / iHeight
  y_size = frm.width / iWidth

  'Loop though all the objects on the form
  'Based on the upper bound of the # of controls
  For i = 0 To UBound(List)
    'Grad each control individually
    For Each curr_obj In frm.Controls
      strTag = UCase(curr_obj.Tag)
      'If TypeOf curr_obj Is Menu Then
    ' ignore menu controls
      'Else
      If strTag <> NORESIZE Then
        If curr_obj.TabIndex = List(i).Index Then
          'Then resize the control
          With curr_obj
            If TypeName(.Container) = "SSTab" Then
                For b = 0 To .Container.Tabs - 1
                    .Container.Tab = b
                    If .Left > 0 Then
                        Exit For
                    End If
                Next
            End If
           
            Select Case strTag
              Case RESIZEALL
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
                .height = List(i).height * x_size
                .Top = List(i).Top * x_size
              Case RESIZELEFT
                .Left = List(i).Left * y_size
              Case RESIZEUPLEFTWIDTH
                .Top = List(i).Top * x_size
                .width = List(i).width * y_size
                .Left = List(i).Left * y_size
              Case RESIZEUPWIDTH
                .Top = List(i).Top * x_size
                .width = List(i).width * y_size
              Case RESIZELEFTWIDTH
                .Left = List(i).Left * y_size
                .width = List(i).width * y_size
              Case RESIZEUPLEFT
                .Left = List(i).Left * y_size
                .Top = List(i).Top * x_size
              Case RESIZEUP
                .Top = List(i).Top * x_size
              Case Else
            End Select
          End With
        End If
      End If
    Next curr_obj
  Next i
End Sub

Public Function SetFontSize() As Integer
  On Error Resume Next
  
    'Make sure x_size is greater than 0
    If Int(x_size) > 0 Then
    'Set the font size
        SetFontSize = Int(x_size * 8)
    End If
End Function

Public Sub GetLocation(frm As Form)
  Dim i, b As Integer
  Dim curr_obj As Control
'   Load the current positions of each object into a user defined type array.
'   This information will be used to rescale them in the Resize function.

'Loop through each control
On Error Resume Next
For Each curr_obj In frm.Controls
'Resize the Array by 1, and preserve
'the original objects in the array
    If UCase(curr_obj.Tag) <> UCase("NORESIZE") Then
        ReDim Preserve List(i)
        With List(i)
            If TypeName(curr_obj.Container) = "SSTab" Then
                For b = 0 To curr_obj.Container.Tabs - 1
                    curr_obj.Container.Tab = b
                    If curr_obj.Left > 0 Then
                        Exit For
                    End If
                Next
            End If
            .Name = curr_obj.Name
            .Index = curr_obj.TabIndex
            .Left = curr_obj.Left
            .Top = curr_obj.Top
            .width = curr_obj.width
            .height = curr_obj.height
        End With
    i = i + 1
    End If
Next curr_obj
    
'   This is what the object sizes will be compared to on rescaling.
    iHeight = frm.height
    iWidth = frm.width
End Sub

Public Sub CenterForm(frm As Form)
  On Error Resume Next
  frm.Move (Screen.width - frm.width) \ 2, (Screen.height - frm.height) \ 2
End Sub

Public Sub ResizeForm(frm As Form)
  On Error Resume Next
    'Set the forms height
    frm.height = Screen.height
    'Set the forms width
    frm.width = Screen.width
    'Resize all of the controls
    'based on the forms new size
    ResizeControls frm
End Sub

Public Sub SetNoResize(ByRef frm As Form) 'this function set noresize to all controls that are not resizeable.Ej: Timers
  Dim ctl As Control
  On Error Resume Next
  
  For Each ctl In frm.Controls
    Select Case UCase(TypeName(ctl))
      Case UCase(Timer)
        ctl.Tag = NORESIZE
      Case UCase(Menu)
        ctl.Tag = NORESIZE
      Case UCase(ImageList)
        ctl.Tag = NORESIZE
      Case UCase(Line)
        ctl.Tag = NORESIZE
    End Select
  Next
End Sub


Public Sub SetResize(ByRef frm As Form, Optional ByVal strResizeType As String, Optional ByVal strCtrlType As String)  'This function sets the default resize to the controls in the form
  Dim ctl As Control
  Dim strReType As String
  On Error Resume Next

  ' you can pass through strCtrlType (strType) these parameters: Typename(object) or just the string with the type of the object you want to apply the changes
  If strResizeType <> "" Then
    strReType = strResizeType
  Else
    strReType = RESIZEDEFAULT
  End If

  If strCtrlType <> "" Then
    For Each ctl In frm.Controls
      Select Case True
        Case TypeName(ctl) = strCtrlType
          ctl.Tag = strReType
      End Select
    Next
  Else
    For Each ctl In frm.Controls
      If strResizeType <> "" Then ' Resize all controls to this type
        ctl.Tag = strReType
      Else
     'Defaults RESIZES
        Select Case UCase(TypeName(ctl))
          Case UCase(CommandButton)
            ctl.Tag = strReType
          Case UCase(ListView)
            ctl.Tag = RESIZEALL
          Case UCase(SSTab)
            ctl.Tag = RESIZEALL
          Case UCase(TextBox)
            ctl.Tag = strReType
          Case UCase(ComboBox)
            ctl.Tag = strReType
          Case UCase(Frame)
            ctl.Tag = RESIZEALL
          Case UCase(CheckBox)
            ctl.Tag = RESIZEUPLEFT
          Case UCase(TreeView)
            ctl.Tag = RESIZEALL
          Case UCase(MaskEdBox)
            ctl.Tag = strReType
          Case UCase(Label)
            If ctl.Caption <> "" Then
              ctl.Tag = RESIZEUPLEFT
            Else
              ctl.Tag = strReType
            End If
          Case UCase(RichTextBox)
            ctl.Tag = RESIZEALL
          Case Else
            ctl.Tag = strReType
        End Select
      End If
    Next
  End If
  SetNoResize frm ' Execute noresize to establish wich controls do not resize

End Sub


