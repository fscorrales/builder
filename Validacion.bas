Attribute VB_Name = "Validacion"
Public Function ValidaRelacion() As Boolean

    If Trim(CargaRelacion.cmbTablaOrigen.Text) = "" Then
        MsgBox "Debe ingresar una Tabla de Origen", vbCritical + vbOKOnly, "TABLA DE ORIGEN INVALIDA"
        CargaRelacion.cmbTablaOrigen.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    If Trim(CargaRelacion.cmbTablaDestino.Text) = "" Then
        MsgBox "Debe ingresar una Tabla de Destino", vbCritical + vbOKOnly, "TABLA DE DESTINO INVALIDA"
        CargaRelacion.cmbTablaDestino.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    If CargaRelacion.cmbTablaOrigen.Text = CargaRelacion.cmbTablaDestino.Text Then
        MsgBox "La tabla Origen y Destino no pueden ser iguales", vbCritical + vbOKOnly, "TABLA DE DESTINO U ORIGEN INVALIDA"
        CargaRelacion.cmbTablaOrigen.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    If Trim(CargaRelacion.cmbCampoOrigen.Text) = "" Then
        MsgBox "Debe ingresar una Campo de Origen", vbCritical + vbOKOnly, "CAMPO DE ORIGEN INVALIDO"
        CargaRelacion.cmbCampoOrigen.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    Dim Indice As Index
    Dim Tabla As TableDef
    Dim ExisteIndice As Boolean
    Set Tabla = dbBuilder.TableDefs(CargaRelacion.cmbTablaOrigen.Text)
    For Each Indice In Tabla.Indexes
        If Right(Indice.Fields, Len(Indice.Fields) - 1) = CargaRelacion.cmbCampoOrigen.Text And Indice.Unique = True Then
            ExisteIndice = True
            Exit For
        End If
    Next Indice
    If ExisteIndice = False Then
        MsgBox "El Campo Origen debe tener un Índice Único", vbCritical + vbOKOnly, "CAMPO DE ORIGEN SIN INDICE UNICO"
        CargaRelacion.cmbCampoOrigen.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    If Trim(CargaRelacion.cmbCampoDestino.Text) = "" Then
        MsgBox "Debe ingresar una Campo de Destino", vbCritical + vbOKOnly, "CAMPO DE DESTINO INVALIDO"
        CargaRelacion.cmbCampoDestino.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    If Trim(CargaRelacion.txtNombreRelacion.Text) = "" Then
        MsgBox "Debe ingresar un Nombre de la Relación", vbCritical + vbOKOnly, "NOMBRE DE LA RELACION INVALIDO"
        CargaRelacion.txtNombreRelacion.SetFocus
        ValidaRelacion = False
        Exit Function
    End If
    Dim Relacion As Relation
    For Each Relacion In dbBuilder.Relations
        If Relacion.Name = CargaRelacion.txtNombreRelacion Then
            MsgBox "El Nombre de la Relacion Ingresado ya existe", vbCritical + vbOKOnly, "NOMBRE DE LA RELACION DUPLICADO"
            CargaRelacion.txtNombreRelacion.SetFocus
            ValidaRelacion = False
            Exit Function
        End If
    Next Relacion

    ValidaRelacion = True

End Function


