Attribute VB_Name = "Conexiones"
Option Compare Text
Public dbBuilder As Database
Public strEditandoTabla As String
Public strEditandoCampo As String
Public rstCargadgRegistros As Recordset
Public rstImportarRegistros As Recordset
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Sub NuevaBD()
    
    Dim Respuesta As Integer
    Dim strDireccion As String
    Principal.dlgMultifuncion.ShowSave
    strDireccion = Principal.dlgMultifuncion.FileName
    If strDireccion = "" Then 'En caso de apretar cancelar
        Exit Sub
    Else
        If Right(strDireccion, 4) Like ".???" Then 'Se ajusta a la extensión .mdb del archivo
            If Right(strDireccion, 4) <> ".mdb" Then
                strDireccion = Left(strDireccion, Len(strDireccion) - 4) & ".mdb"
            End If
        Else
            strDireccion = Principal.dlgMultifuncion.FileName & ".mdb"
        End If
    End If
    
    If Principal.Caption <> "BUILDER" Then
        dbBuilder.Close
    End If
    
    If FileSystem.Dir(strDireccion) <> "" Then ' Si la BD ya existe
        Respuesta = MsgBox("Una base de datos ya Existe, esta seguro que desea REEMPLAZARLA?", 4, "Alerta")
        If Respuesta = 6 Then 'Al aceptar
            FileSystem.Kill (strDireccion)
        Else 'Al apretar Cancelar
            MsgBox "Se canceló la operación", 0 + 48
            strDireccion = ""
            Exit Sub
        End If
    End If

    'Preparando Base de Datos
    Dim dbBase As Database

    'Preparando Espacio de Trabajo
    Dim wsEspacio As Workspace

    'Activando el Espacio de Trabajo
    Set wsEspacio = DBEngine.Workspaces(0)

    'Generando la Base de Datos
    Set dbBase = wsEspacio.CreateDatabase(strDireccion, dbLangGeneral, dbVersion30)
    
    dbBase.Close
    wsEspacio.Close
    Conectar (strDireccion)
    strDireccion = ""
    
End Sub

Public Sub Conectar(Direccion As String)
    
    'En caso de venir desde NuevaDB
    If Direccion <> "" Then
        Set dbBuilder = OpenDatabase(Direccion)
        Principal.Caption = "BUILDER Conected to " & Direccion
        Conectado (True)
        Exit Sub
    End If
    
    'En los demás casos
    Dim strDireccion As String
    Principal.dlgMultifuncion.Filter = "Todos los Access (*.mdb)|*.mdb|"
    Principal.dlgMultifuncion.FileName = ""
    Principal.dlgMultifuncion.ShowOpen
    strDireccion = Principal.dlgMultifuncion.FileName
    If strDireccion = "" Then 'En caso de apretar cancelar
        Exit Sub
    Else
        Conectar (strDireccion)
        strDireccion = ""
    End If
    
End Sub

Public Sub Conectado(ExisteConexión As Boolean)

    If ExisteConexión = False Then
        If Principal.Caption <> "BUILDER" Then
            dbBuilder.Close
        End If
        With Principal
            .Caption = "BUILDER"
            .MnuDesconectar.Enabled = False
            .MnuNuevaTabla.Enabled = False
            .MnuListadoTablas.Enabled = False
            .MnuNuevoCampo.Enabled = False
            .MnuListadoCampos.Enabled = False
            .MnuIndice.Enabled = False
            .MnuNuevaRelacion.Enabled = False
            .MnuListadoRelaciones.Enabled = False
            .MnuListadoRegistros.Enabled = False
            .MnuImportarXLS.Enabled = False
        End With
    Else
        With Principal
            .MnuDesconectar.Enabled = True
            .MnuNuevaTabla.Enabled = True
            .MnuListadoTablas.Enabled = True
            .MnuNuevoCampo.Enabled = True
            .MnuListadoCampos.Enabled = True
            .MnuIndice.Enabled = True
            .MnuNuevaRelacion.Enabled = True
            .MnuListadoRelaciones.Enabled = True
            .MnuListadoRegistros.Enabled = True
            .MnuImportarXLS.Enabled = True
        End With
    End If
    
End Sub

Public Sub SetRecordset(NombreRecordset As Recordset, SQL As String)
    
    Set NombreRecordset = dbBuilder.OpenRecordset(SQL)

End Sub

Public Sub GenerarTabla(NombreTabla As String)

    Dim Tabla As TableDef
    If NombreTabla <> "" Then
        Set Tabla = dbBuilder.CreateTableDef(NombreTabla)
        With Tabla
            .Fields.Append .CreateField("CampoProvisorio", dbText, 1)
        End With
        dbBuilder.TableDefs.Append Tabla
        Unload CargaTabla
        ConfigurardgTablas
        CargardgTablas
    End If
    
End Sub

Public Sub EditarTabla()

    If strEditandoTabla = "" Then
        Dim i As String
        i = ListadoTablas.dgTablas.Row
        If i <> 0 And i <> ListadoTablas.dgTablas.Rows - 1 Then
            strEditandoTabla = ListadoTablas.dgTablas.TextMatrix(i, 0)
            With CargaTabla
                .Show
                .Frame1.Caption = "Editar Tabla"
                .txtTabla = strEditandoTabla
            End With
            Unload ListadoTablas
        End If
        i = ""
    Else
        Dim Tabla As TableDef
        Set Tabla = dbBuilder.TableDefs(strEditandoTabla)
        Tabla.Name = CargaTabla.txtTabla.Text
        Unload CargaTabla
        ListadoTablas.Show
        ConfigurardgTablas
        CargardgTablas
    End If

End Sub

Public Sub EliminarTabla()

    Dim Tabla As TableDef
    Dim i As String
    Dim Borrar As Integer
    i = ListadoTablas.dgTablas.Row
    If i <> 0 And i <> ListadoTablas.dgTablas.Rows - 1 Then
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la TABLA: " & ListadoTablas.dgTablas.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que todos los registros contenidos y ascociados a la misma también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO TABLA")
        If Borrar = 6 Then
            dbBuilder.TableDefs.Delete ListadoTablas.dgTablas.TextMatrix(i, 0)
            CargardgTablas
        End If
        Borrar = 0
    End If
    i = ""
    
End Sub

Public Sub GenerarCampo()

    Dim Tabla As TableDef
    Dim Campo As Field
    
    Set Tabla = dbBuilder.TableDefs(CargaCampo.cmbNombreTabla.Text)
    
    With Tabla
        Select Case CargaCampo.cmbTipo.Text
        Case "Texto"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbText, CargaCampo.txtTamaño.Text)
        Case "Moneda"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbCurrency)
        Case "Long"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbLong)
        Case "Integer"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbInteger)
        Case "Byte"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbByte)
        Case "Date/Time"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbDate)
        Case "Boleano"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbBoolean)
        Case "Single"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbSingle)
        Case "Double"
            .Fields.Append .CreateField(CargaCampo.txtNombreCampo.Text, dbDouble)
        End Select
    End With
    
    With Tabla.Fields(CargaCampo.txtNombreCampo.Text)
        If CargaCampo.chkLongitud.Value = 1 And CargaCampo.cmbTipo.Text = "Texto" Then
            .AllowZeroLength = True
        Else
            .AllowZeroLength = False
        End If
        If CargaCampo.chkRequerido.Value = 1 Then
            .Required = True
        Else
            .Required = False
        End If
    End With

    For Each Campo In Tabla.Fields
        If Campo.Name = "CampoProvisorio" Then
            Tabla.Fields.Delete "CampoProvisorio"
            Exit For
        End If
    Next
    
    ListadoCampos.Show
    ListadoCampos.cmbNombreTabla.Clear
    Call CargarcmbNombreTabla(ListadoCampos)
    ListadoCampos.cmbNombreTabla.Text = CargaCampo.cmbNombreTabla.Text
    ConfigurardgCampos
    Call CargardgCampos(ListadoCampos.cmbNombreTabla.Text)
    Unload CargaCampo
    
End Sub

Public Sub EliminarCampo()

    Dim Tabla As TableDef
    Dim Campo As Field
    Dim i As String
    Dim Borrar As Integer
    i = ListadoCampos.dgCampos.Row
    If i <> 0 Then
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el CAMPO: " & ListadoCampos.dgCampos.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que todos los registros contenidos y ascociados al mismo también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO CAMPO")
        If Borrar = 6 Then
            Set Tabla = dbBuilder.TableDefs(ListadoCampos.cmbNombreTabla.Text)
            Tabla.Fields.Delete ListadoCampos.dgCampos.TextMatrix(i, 0)
            Call CargardgCampos(ListadoCampos.cmbNombreTabla.Text)
        End If
        Borrar = 0
    End If
    i = ""
    
End Sub

Public Sub EditarCampo()

    If strEditandoCampo = "" Then
        Dim i As String
        i = ListadoCampos.dgCampos.Row
        If i <> 0 Then
            strEditandoCampo = ListadoCampos.dgCampos.TextMatrix(i, 0)
            With CargaCampo
                .Show
                Call CargarcmbNombreTabla(CargaCampo)
                CargarcmbTipoCampo
                .Frame1.Caption = "Tabla a la que pertenece el Campo a EDITAR"
                .cmbNombreTabla.Text = ListadoCampos.cmbNombreTabla.Text
                .cmbNombreTabla.Enabled = False
                .cmbTipo.Text = ListadoCampos.dgCampos.TextMatrix(i, 2)
                .txtTamaño.Text = ListadoCampos.dgCampos.TextMatrix(i, 3)
                If ListadoCampos.dgCampos.TextMatrix(i, 4) = True Then
                    .chkRequerido.Value = 1
                End If
                If ListadoCampos.dgCampos.TextMatrix(i, 5) = True Then
                    .chkLongitud.Value = 1
                End If
                .txtNombreCampo.Text = strEditandoCampo
            End With
            Unload ListadoCampos
        End If
        i = ""
    Else
        Dim strTabla As String
        Dim strCampo As String
        Dim qdf As QueryDef
        Dim Campo As Field
        Dim Tabla As TableDef
        
        strTabla = CargaCampo.cmbNombreTabla.Text
        strCampo = CargaCampo.txtNombreCampo.Text
        'Create a dummy QueryDef object.
        Set qdf = dbBuilder.CreateQueryDef("", "Select * from " & strTabla)
        'Add a temporary field to the table.
        Select Case CargaCampo.cmbTipo.Text
        Case "Texto"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField text (" & CargaCampo.txtTamaño.Text & ")"
        Case "Moneda"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField currency"
        Case "Long"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField long"
        Case "Integer"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField integer"
        Case "Byte"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField byte"
        Case "Date/Time"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField date"
        Case "Boleano"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField boolean"
        Case "Double"
            qdf.SQL = "ALTER TABLE [" & strTabla & "] ADD COLUMN AlterTempField Double"
        End Select
        qdf.Execute
        'Copy the data from old field into the new field.
        qdf.SQL = "UPDATE DISTINCTROW [" & strTabla & "] SET  AlterTempField = [" & strEditandoCampo & "]"
        qdf.Execute
        'Agregar un 0 a todos los registrso de un campo
        'qdf.SQL = "UPDATE [" & strTabla & "] SET AlterTempField = '0' & AlterTempField"
        'qdf.Execute
        'Delete the old field.
        qdf.SQL = "ALTER TABLE [" & strTabla & "] DROP COLUMN [" & strEditandoCampo & "]"
        qdf.Execute
        
        'Rename the temporary field to the old field's name.
        Set Tabla = dbBuilder.TableDefs(strTabla)
        Tabla.Fields.Refresh
        Set Campo = Tabla.Fields("AlterTempField")
        Campo.Name = strCampo
         
        'Set Tabla = dbBuilder.TableDefs(CargaCampo.cmbNombreTabla.Text)
        'Set Campo = Tabla.Fields(strEditandoCampo)
        With Campo
            If CargaCampo.chkLongitud.Value = 1 And CargaCampo.cmbTipo.Text = "Texto" Then
                .AllowZeroLength = True
            Else
                .AllowZeroLength = False
            End If
            If CargaCampo.chkRequerido.Value = 1 Then
                .Required = True
            Else
                .Required = False
            End If
        End With
        
        ListadoCampos.Show
        ListadoCampos.cmbNombreTabla.Clear
        Call CargarcmbNombreTabla(ListadoCampos)
        ListadoCampos.cmbNombreTabla.Text = CargaCampo.cmbNombreTabla.Text
        ConfigurardgCampos
        Call CargardgCampos(ListadoCampos.cmbNombreTabla.Text)
        Unload CargaCampo
    End If

End Sub


Public Sub GenerarIndice()

    Dim Tabla As TableDef
    Dim Indice As Index
    Dim i As Integer
    Dim Denominacion As String
    
    Denominacion = CargaIndice.cmbNombreTabla.Text
    Set Tabla = dbBuilder.TableDefs(Denominacion)
    
    If Not Tabla.Indexes.Count = 0 Then
        For i = 0 To Tabla.Indexes.Count - 1
            Set Indice = Tabla.Indexes(i)
            If Indice.Primary = True Then
                Tabla.Indexes.Delete (Indice.Name)
            End If
        Next i
    End If
    
    If Not CargaIndice.lstCampoIndice.ListCount = 0 Then
        Set Indice = Tabla.CreateIndex("pk" & Denominacion)
        For i = 0 To CargaIndice.lstCampoIndice.ListCount - 1
            CargaIndice.lstCampoIndice.ListIndex = i
            Denominacion = CargaIndice.lstCampoIndice.Text
            Indice.Fields.Append Indice.CreateField(Denominacion)
        Next i
        With Indice
            '.Fields.Append .CreateField(denominacion)
            .Primary = True
            .Unique = True
        End With
        Tabla.Indexes.Append Indice
    End If
    
    ListadoCampos.Show
    ConfigurardgCampos
    Call CargarcmbNombreTabla(ListadoCampos)
    ListadoCampos.cmbNombreTabla.Text = CargaIndice.cmbNombreTabla.Text
    CargardgCampos (CargaIndice.cmbNombreTabla.Text)
    Unload CargaIndice
    
End Sub

Public Sub GenerarRelacion()

    'Definición de las relaciones entre tablas empleando el objeto relación
    Dim Relacion As Relation
    Dim CampoRel As Field

    Set Relacion = dbBuilder.CreateRelation(CargaRelacion.txtNombreRelacion.Text, CargaRelacion.cmbTablaOrigen.Text, _
    CargaRelacion.cmbTablaDestino.Text, dbRelationUpdateCascade + dbRelationDeleteCascade)
    'Crear el campo de relación y establecer las propiedades
    Set CampoRel = Relacion.CreateField(CargaRelacion.cmbCampoOrigen.Text)
    CampoRel.ForeignName = CargaRelacion.cmbCampoDestino.Text
    'Agregar el campo a la relación y la relación a la DB
    Relacion.Fields.Append CampoRel
    dbBuilder.Relations.Append Relacion

    Unload CargaRelacion
    ListadoRelaciones.Show
    ConfigurardgRelaciones
    CargardgRelaciones
    
End Sub

Public Sub EliminarRelacion()

    Dim i As String
    Dim Borrar As Integer
    i = ListadoRelaciones.dgRelaciones.Row
    If i <> 0 Then
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la RELACION: " & ListadoRelaciones.dgRelaciones.TextMatrix(i, 4) & "?", vbQuestion + vbYesNo, "BORRANDO RELACION")
        If Borrar = 6 Then
            dbBuilder.Relations.Delete ListadoRelaciones.dgRelaciones.TextMatrix(i, 4)
            CargardgRelaciones
        End If
        Borrar = 0
    End If
    i = ""
    
End Sub

Public Sub DefinirRutaAcceso()

    Dim strDireccion As String
    Principal.dlgMultifuncion.Filter = "Todos los Excel (*.xls)|*.xls|"
    Principal.dlgMultifuncion.FileName = ""
    Principal.dlgMultifuncion.ShowOpen
    strDireccion = Principal.dlgMultifuncion.FileName
    If strDireccion = "" Then 'En caso de apretar cancelar
        Exit Sub
    Else
        ImportarXLS.txtDireccion.Text = strDireccion
        strDireccion = ""
    End If

End Sub
