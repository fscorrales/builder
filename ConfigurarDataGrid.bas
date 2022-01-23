Attribute VB_Name = "ConfigurarDataGrid"
Public Sub ConfigurardgTablas()
    
    With ListadoTablas.dgTablas
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Nombre"
        .ColWidth(0) = 2000
        .TextMatrix(0, 1) = "N° Campos"
        .ColWidth(1) = 1000
        .TextMatrix(0, 2) = "N° Registros"
        .ColWidth(2) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 7
    End With
    
End Sub
Public Sub CargardgTablas()
    
    Dim Tabla As TableDef
    Dim i As Integer
    Dim lngContarCampos As Long
    Dim lngContarRegistros As Long
    Dim Campo As Field
    i = 0
    ListadoTablas.dgTablas.Rows = 2
    
    
    
    For Each Tabla In dbBuilder.TableDefs
        If Not Tabla.Name Like "MSys*" Then
            i = i + 1
            With ListadoTablas.dgTablas
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = Tabla.Name
                For Each Campo In Tabla.Fields
                    If Campo.Name = "CampoProvisorio" Then
                        .TextMatrix(i, 1) = Tabla.Fields.Count - 1
                        Exit For
                    Else
                        .TextMatrix(i, 1) = Tabla.Fields.Count
                    End If
                Next
                lngContarCampos = lngContarCampos + .TextMatrix(i, 1)
                .TextMatrix(i, 2) = Tabla.RecordCount
                lngContarRegistros = lngContarRegistros + Tabla.RecordCount
                .Rows = .Rows + 1
            End With
        End If
    Next
    i = i + 1
    With ListadoTablas.dgTablas
        .RowHeight(i) = 300
        .TextMatrix(i, 0) = "Totales"
        .TextMatrix(i, 1) = lngContarCampos
        .TextMatrix(i, 2) = lngContarRegistros
    End With

End Sub

Public Sub ConfigurardgCampos()
    
    With ListadoCampos.dgCampos
        .Clear
        .Cols = 6
        .Rows = 2
        .TextMatrix(0, 0) = "Nombre"
        .ColWidth(0) = 2000
        .TextMatrix(0, 1) = "Index"
        .ColWidth(1) = 1000
        .TextMatrix(0, 2) = "Tipo"
        .ColWidth(2) = 1000
        .TextMatrix(0, 3) = "Tamaño"
        .ColWidth(3) = 1000
        .TextMatrix(0, 4) = "Requerido"
        .ColWidth(4) = 1000
        .TextMatrix(0, 5) = "Long. Cero"
        .ColWidth(5) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColAlignment(3) = 7
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
    End With
    
End Sub

Public Sub CargardgCampos(NombreTabla As String)
    
    Dim Tabla As TableDef
    Dim i As Integer
    Dim Campo As Field
    Dim Indice As Index
    i = 0
    ListadoCampos.dgCampos.Rows = 2
    
    Set Tabla = dbBuilder.TableDefs(NombreTabla)
    For Each Campo In Tabla.Fields
        If Not Campo.Name = "CampoProvisorio" Then
            i = i + 1
            With ListadoCampos.dgCampos
                .RowHeight(i) = 300
                .TextMatrix(i, 0) = Campo.Name
                .TextMatrix(i, 1) = False
                For Each Indice In Tabla.Indexes
                    If Indice.Primary = True Then
                        Dim CampoIndezado As Field
                        For Each CampoIndezado In Indice.Fields
                            If CampoIndezado.Name = Campo.Name Then
                                .TextMatrix(i, 1) = True
                                Exit For
                            Else
                                .TextMatrix(i, 1) = False
                            End If
                        Next CampoIndezado
                    End If
                Next Indice
                Select Case Campo.Type
                Case 10
                    .TextMatrix(i, 2) = "Texto"
                Case 5
                    .TextMatrix(i, 2) = "Moneda"
                Case 4
                    .TextMatrix(i, 2) = "Long"
                Case 3
                    .TextMatrix(i, 2) = "Integer"
                Case 2
                    .TextMatrix(i, 2) = "Byte"
                Case 8
                    .TextMatrix(i, 2) = "Date/Time"
                Case 1
                    .TextMatrix(i, 2) = "Boleano"
                Case 6
                    .TextMatrix(i, 2) = "Single"
                Case 7
                    .TextMatrix(i, 2) = "Double"
                Case Else
                    .TextMatrix(i, 2) = "No Definido"
                End Select
                .TextMatrix(i, 3) = Campo.Size
                .TextMatrix(i, 4) = Campo.Required
                .TextMatrix(i, 5) = Campo.AllowZeroLength
                .Rows = .Rows + 1
            End With
        End If
    Next
    ListadoCampos.dgCampos.Rows = ListadoCampos.dgCampos.Rows - 1

End Sub

Public Sub ConfigurardgRelaciones()
    
    With ListadoRelaciones.dgRelaciones
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Tabla Origen"
        .ColWidth(0) = 1500
        .TextMatrix(0, 1) = "Tabla Destino"
        .ColWidth(1) = 1500
        .TextMatrix(0, 2) = "Campo Origen"
        .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "Campo Destino"
        .ColWidth(3) = 1500
        .TextMatrix(0, 4) = "Nombre Relación"
        .ColWidth(4) = 2200
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
    End With
    
End Sub

Public Sub CargardgRelaciones()
    
    Dim Tabla As TableDef
    Dim i As Integer
    Dim Campo As Field
    Dim Relacion As Relation
    i = 0
    ListadoRelaciones.dgRelaciones.Rows = 2
    
    For Each Relacion In dbBuilder.Relations
        i = i + 1
        With ListadoRelaciones.dgRelaciones
            .RowHeight(i) = 300
            .TextMatrix(i, 0) = Relacion.Table
            .TextMatrix(i, 1) = Relacion.ForeignTable
            For Each Campo In Relacion.Fields
                .TextMatrix(i, 2) = Campo.Name
                .TextMatrix(i, 3) = Campo.ForeignName
            Next Campo
            .TextMatrix(i, 4) = Relacion.Name
            .Rows = .Rows + 1
        End With
    Next Relacion
    ListadoRelaciones.dgRelaciones.Rows = ListadoRelaciones.dgRelaciones.Rows - 1

End Sub

Sub ConfigurardgRegistros()

    Dim Tabla As TableDef
    Dim Campo As Field
    Dim i As Integer
    i = 0
    
    Set Tabla = dbBuilder.TableDefs(ListadoRegistros.cmbNombreTabla.Text)
    With ListadoRegistros.dgRegistros
        .Clear
        .Cols = Tabla.Fields.Count
        .Rows = 2
        For Each Campo In Tabla.Fields
            .TextMatrix(0, i) = Campo.Name
            .ColWidth(i) = 2000
            .ColAlignment(i) = 1
            i = i + 1
        Next Campo
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
    End With
End Sub

Sub CargardgRegistros()
    Dim i As Long
    i = 0
    Dim x As Long
    x = 0
    ListadoRegistros.dgRegistros.Rows = 2
    Set Tabla = dbBuilder.TableDefs(ListadoRegistros.cmbNombreTabla.Text)
    Call SetRecordset(rstCargadgRegistros, "SELECT * FROM " & ListadoRegistros.cmbNombreTabla.Text & " ORDER BY " & ListadoRegistros.cmbOrdenCampo.Text)
    If rstCargadgRegistros.BOF = False Then
        With rstCargadgRegistros
            .MoveFirst
            While .EOF = False
                i = i + 1
                ListadoRegistros.dgRegistros.RowHeight(i) = 300
                For x = 0 To Tabla.Fields.Count - 1
                    If Len(.Fields(x)) <> "0" Then
                        ListadoRegistros.dgRegistros.TextMatrix(i, x) = .Fields(x)
                    Else
                        ListadoRegistros.dgRegistros.TextMatrix(i, x) = "0"
                    End If
                Next x
                .MoveNext
                ListadoRegistros.dgRegistros.Rows = ListadoRegistros.dgRegistros.Rows + 1
            Wend
        End With
        ListadoRegistros.dgRegistros.Rows = ListadoRegistros.dgRegistros.Rows - 1
    End If
    ListadoRegistros.dgRegistros.SetFocus

End Sub

