Attribute VB_Name = "Connection"
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As String) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
' The above functions are used for the menu images
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Public Sub connected()
If con.State Then con.Close
con.Open "provider=Microsoft.jet.OLEDB.3.51;Data Source=" & App.Path & "\Hospital.mdb"
'rs.LockType = adLockOptimistic
'rs.CursorType = adOpenDynamic
End Sub

