VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAdo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private adoMome As ADODB.Recordset
Private adoTabla As ADODB.Recordset
Private conMome As ADODB.Connection
Private tTabla As String


Public Sub msSalvaUnAdo(aPrm As ADODB.Recordset, _
    cPrm As ADODB.Connection, tPrm As String)
    'aPrm el recordset que se va a guardar
    'tPrm el nombre de la tabla
    'cPrm la conexion de la tabla

Set adoMome = New ADODB.Recordset
'Set conMome = New ADODB.Connection
Set conMome = cPrm
tTabla = tPrm
Set adoMome = aPrm

'verifica que no este vacio
If adoMome.RecordCount < 1 Then
    Exit Sub
End If
Debug.Print "No esta vacio"

'verifica que existe la tabla
If Not ExisteLaTabla Then
    CreaLaTabla             '// Ojo : revisar
End If

'abre la tabla
Set adoTabla = New ADODB.Recordset
adoTabla.Open "select * from " & tTabla, conMome, adOpenKeyset, adLockOptimistic, adCmdText

'recorre el recordset
Dim nc As Integer
nc = adoMome.Fields.Count
adoMome.MoveFirst
Do While Not adoMome.EOF
    GrabaElRegistro (nc)
    adoMome.MoveNext
Loop
End Sub

Private Sub GrabaElRegistro(nPrm As Integer)
Dim mM As Integer
adoTabla.AddNew
For mM = 0 To nPrm - 1
    adoTabla.Fields(mM) = adoMome.Fields(mM)
Next
adoTabla.Update
End Sub
Private Function ExisteLaTabla() As Boolean
'y supone que tiene la misma estructura que el recordset
'ojo: proyecto,referencia: Microsoft Ado ext
Dim cat As New ADOX.Catalog
Dim nM, nP As Integer

Set cat = New ADOX.Catalog
Set cat.ActiveConnection = conMome
nM = cat.Tables.Count
For nP = 0 To nM - 1
    If cat.Tables(nP).Name = tTabla Then
        Debug.Print "encontro la tabla " & tTabla
        ExisteLaTabla = True
        GoTo fin
        Exit Function
    End If
Next nP
Debug.Print "No encontro la tabla " & tTabla
ExisteLaTabla = False
fin:

Set cat = Nothing
End Function

Private Sub CreaLaTabla()
'ojo esta rutina no funciona bien
Dim cat As New ADOX.Catalog
Dim nM, nP As Integer
Dim tbl As New Table

Set cat = New ADOX.Catalog
Set cat.ActiveConnection = adoconn
Set tbl = New Table
tbl.Name = tTabla
nM = adoMome.Fields.Count
'For np = 0 To nm - 1
'    tbl.Columns.Append adoMome.Fields(np).Name, _
'        adoMome.Fields(np).DefinedSize, _
'        adoMome.Fields(np).Type
'Next np
'End Sub
tbl.Columns.Append "prueba", adBoolean, 0

cat.Tables.Append tbl
'Set cat = Nothing
End Sub
