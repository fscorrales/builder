Attribute VB_Name = "Importacion"
Public Sub ImportarRegistrosXLS()

    Dim strDireccionExcel As String
    Dim strDireccionDB As String
    Dim strTablaOrigen As String
    Dim strTablaDestino As String
    Dim strConnect As String, strSQL As String
    Dim dbImportar As ADODB.Connection
    Dim rstImportar As ADODB.Recordset
    Dim RecsAffected As Long
    
    With ImportarXLS
        strDireccionExcel = .txtDireccion.Text
        strDireccionDB = Right(Principal.Caption, Len(Principal.Caption) - 20)
        strTablaOrigen = "[Hoja1$" & .txtCeldaInicio.Text & ":" & .txtCeldaFin.Text & "]"
        strTablaDestino = .cmbTabla.Text
    End With

    'Debug.Print strDireccionExcel
    'Debug.Print strDireccionDB
    'Debug.Print strTablaOrigen
    'Debug.Print strTablaDestino
    
    ' Establezco la conexión con la base de datos de Access,
    ' la cual será la base de datos "Activa"

    Set dbImportar = New ADODB.Connection
    dbImportar.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & strDireccionExcel & ";" & _
    "Extended Properties='Excel 8.0;HDR=Yes'"
    'Set rstImportar = New ADODB.Recordset
    'rstImportar.Open strTablaDestino, dbImportar, adOpenDynamic, adLockOptimistic
    
    'Rango que quiero importar dela hoja Sheet1
    'TablaOrigen = "[Sheet1$A1:C1500]"

    ' Importo la tabla a la base de datos "Activa"
    strConnect = "'" & strDireccionExcel & "' 'Excel 8.0;HDR=Yes;'"

    strSQL = "INSERT INTO " & strTablaDestino & " In '" & _
    strDireccionDB & _
    "' SELECT * FROM " & strTablaOrigen
'    Debug.Print strSQL
    dbImportar.Execute strSQL, RecsAffected, adCmdText + adExecuteNoRecords

    'Debug.Print RecsAffected

    ' Cierro la conexión
    dbImportar.Close
    Set dbImportar = Nothing
    Unload ImportarXLS
    MsgBox "Se importaron correctamente " & RecsAffected & " registros", vbInformation + vbOKOnly, "IMPORTACION COMPLETADA"
End Sub

