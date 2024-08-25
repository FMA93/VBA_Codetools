Attribute VB_Name = "MinimiseMaximiseIcons"
Option Explicit

#If VBA7 Then
    ' 64-bit Excel
    #If Win64 Then
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
        Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
        Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    #Else
        ' 32-bit Excel on 64-bit Windows
        Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
        Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
        Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    #End If
#Else
    ' 32-bit Excel
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
#End If

Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Public Sub AddMaximizeButton(ByVal UserForm As Object)
    Dim hWnd As LongPtr, lStyle As LongPtr
    hWnd = FindWindow("ThunderDFrame", UserForm.Caption)
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    lStyle = lStyle Or WS_MAXIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, lStyle
End Sub

Public Sub AddMinimizeButton(ByVal UserForm As Object)
    Dim hWnd As LongPtr, lStyle As LongPtr
    hWnd = FindWindow("ThunderDFrame", UserForm.Caption)
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX
    SetWindowLong hWnd, GWL_STYLE, lStyle
End Sub

#If VBA7 Then
    ' 64-bit Excel
    #If Win64 Then
        Private Function GetWindowLong(ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
            GetWindowLong = GetWindowLongPtr(hWnd, nIndex)
        End Function
        Private Function SetWindowLong(ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
            SetWindowLong = SetWindowLongPtr(hWnd, nIndex, dwNewLong)
        End Function
    #End If
#End If

