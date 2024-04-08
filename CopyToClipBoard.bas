Attribute VB_Name = "CopyToClipBoard"
Option Explicit

Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
   ByVal dwBytes As Long) As Long

Private Declare Function CloseClipboard Lib "User32" () As Long

Private Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long

Private Declare Function EmptyClipboard Lib "User32" () As Long

Private Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, _
   ByVal lpString2 As Any) As Long

Private Declare Function SetClipboardData Lib "User32" (ByVal wFormat _
   As Long, ByVal hMem As Long) As Long

Private Const GHND = &H42
Private Const CF_TEXT = 1
Private Const MAXSIZE = 4096
Sub CopyEntityNameToClipbord(tableName As String, rsTableIds As Recordset, con As Connection)
    Dim rs As Recordset
    Dim sql As String
    Dim res As String
    Dim tableId As String
    
    While Not rsTableIds.EOF
        tableId = rsTableIds(0)
        sql = " select name from lims_sys." & tableName
        sql = sql & " where  " & tableName & "_id= " & tableId
        Set rs = con.Execute(sql)
        
        If Not rs.EOF Then
            res = res & rs(0) & vbCrLf
        End If
        rsTableIds.MoveNext
    Wend
    res = Left(res, Len(res) - 2)
    ClipBoard_SetData (res)
End Sub
Public Function ClipBoard_SetData(MyString As String)
   Dim hGlobalMemory As Long, lpGlobalMemory As Long
   Dim hClipMemory As Long, X As Long

   ' Allocate moveable global memory.
   '-------------------------------------------
   hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

   ' Lock the block to get a far pointer
   ' to this memory.
   lpGlobalMemory = GlobalLock(hGlobalMemory)

   ' Copy the string to this global memory.
   lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

   ' Unlock the memory.
   If GlobalUnlock(hGlobalMemory) <> 0 Then
      MsgBox "Could not unlock memory location. Copy aborted."
      GoTo OutOfHere2
   End If

   ' Open the Clipboard to copy data to.
   If OpenClipboard(0&) = 0 Then
      MsgBox "Could not open the Clipboard. Copy aborted."
      Exit Function
   End If

   ' Clear the Clipboard.
   X = EmptyClipboard()

   ' Copy the data to the Clipboard.
   hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

   If CloseClipboard() = 0 Then
      MsgBox "Could not close Clipboard."
   End If
End Function



