Imports System.Data.SqlClient
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    Dim cadenaConexion As String = "Server=10.195.10.166,1433;Database=ScanSystemDB;User Id=Manu;Password=2022.Tgram2;"
    Private ContadorPiezas As Integer = 0
    Private UltimoPeso As Double = 0
    Private VarPesoMin As Double = 0.03
    Private VarPesoMax As Double = 0.5
    Private PesoEsperado As Double = 0
    Private EstadoBloqueo As Boolean = False
    Private empleadoOk As Boolean = False
    Private mandrilOk As Boolean = False
    Private mesaOk As Boolean = False
    Private ultimoMandril As String = ""
    Private WithEvents bascula As New BasculaReader
    ' Variable global para el 
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InicializarPantalla()
    End Sub
    Private Sub InicializarPantalla()
        ' --- Limpiar DataGridViews ---
        DGVPiezasBuenas.DataSource = Nothing
        DGVPiezasBuenas.Rows.Clear()
        DGVDefectos.DataSource = Nothing
        DGVDefectos.Rows.Clear()
        ' --- Reiniciar Labels ---
        LabelBuenas.Text = "Buenas: 0"
        LabelDefectos.Text = "Defectuosas: 0"
        LabelTotal.Text = "Total: 0"
        ' --- Reiniciar mandril (variable y Label) ---
        ultimoMandril = String.Empty
        LabelMandril.Text = "Mandril: -"
        LabelSP.Text = 0
        ' --- Resetear otros controles si aplica ---
        LabelMesa.Text = ""
        LabelNETM.Text = ""
        LabelNameTM.Text = "Escanear Numero de Empleado"
        ' --- Opcional: limpiar TextBox, ComboBox, etc. ---
        TextBoxInput.Clear()
        Mesa()
        Timer1.Interval = 3000
        Timer1.Start()
        TextBoxInput.Focus()
        Mayusculas()
        CargarMandrilesDistribucion()
        ' Validar conexión al iniciar
        '  VerificarConexionBascula()
    End Sub

    Private Function ObtenerMesaDesdeIni() As String
        Dim configPath As String = "config.ini"
        Dim mesaIdArchivo As String = ""
        If File.Exists(configPath) Then
            Dim lines() As String = File.ReadAllLines(configPath)
            For Each line As String In lines
                line = line.Trim()
                If line.StartsWith("Mesa_Id") Then
                    mesaIdArchivo = line.Split("="c)(1).Trim()
                End If
            Next
        End If
        Return mesaIdArchivo
    End Function

    Private Sub CargarMandrilesDistribucion()
        Dim mesaIni As String = ObtenerMesaDesdeIni()
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            ' Traer solo columna Mandril, filtrando por Estacion (igual a Mesa_Id) y Area = 'Inspeccion'
            Dim query As String = "SELECT Mandril
                               FROM Mandriles
                               WHERE Estacion = @mesa AND Area = 'Inspeccion'"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@mesa", mesaIni)
                Dim adapter As New SqlDataAdapter(cmd)
                Dim tabla As New DataTable()
                adapter.Fill(tabla)
                DGVDistribucion.DataSource = tabla
                DGVDistribucion.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                DGVDistribucion.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End Using
        End Using
    End Sub

    ' --- Sub para validar y reconectar la báscula ---
    Private Sub VerificarConexionBascula()
        Try
            If bascula Is Nothing Then
                bascula = New BasculaReader()
            End If

            ' Intentar iniciar si no está conectada
            If Not bascula.EstaConectada() Then
                If bascula.Iniciar() Then
                    LabelAyuda.Text = "✅ Báscula conectada correctamente"
                    LabelAyuda.BackColor = Color.LightGreen
                Else
                    LabelAyuda.Text = "⚠ Báscula no conectada ⚠"
                    LabelAyuda.BackColor = Color.Red
                End If
            Else
                LabelAyuda.Text = "✅ Báscula conectada correctamente"
                LabelAyuda.BackColor = Color.LightGreen
            End If

        Catch ex As Exception
            LabelAyuda.Text = "⚠ Error al conectar con la báscula ⚠"
            LabelAyuda.BackColor = Color.Red
        End Try
    End Sub
    Private Sub ProcesarPesoBascula(peso As Double, crudo As String)
        If Me.InvokeRequired Then
            Me.BeginInvoke(Sub() ProcesarPesoBascula(peso, crudo))
            Return
        End If

        Dim diferencia As Double = peso - UltimoPeso

        ' Filtro de ruido
        If Math.Abs(diferencia) < VarPesoMin Then
            UltimoPeso = peso
            Exit Sub
        End If

        ' Bloqueo activo
        If EstadoBloqueo Then
            UltimoPeso = peso
            Exit Sub
        End If

        ' --- Incremento ---
        If diferencia > 0 Then
            If diferencia > VarPesoMax Then
                LabelAyuda.Text = "⚠ Peso fuera de tolerancia ⚠"
                LabelAyuda.BackColor = Color.Yellow
                UltimoPeso = peso
                Exit Sub
            End If

            ContadorPiezas += 1
            LabelContador.Text = ContadorPiezas.ToString()
            LabelAyuda.Text = "📦 Se colocó una pieza"
            LabelAyuda.BackColor = Color.FromArgb(127, 179, 131)
            PesoEsperado = peso

            ' --- Decremento ---
        Else
            Dim caida As Double = Math.Abs(diferencia)
            If caida > VarPesoMax Then
                LabelAyuda.Text = "⚠ Se retiraron varias piezas ⚠"
                LabelAyuda.BackColor = Color.Yellow
                UltimoPeso = peso
                Exit Sub
            End If

            If ContadorPiezas > 0 Then
                ContadorPiezas -= 1
                LabelContador.Text = ContadorPiezas.ToString()
                LabelAyuda.Text = "📦 Se retiró una pieza"
                LabelAyuda.BackColor = Color.LightBlue
                PesoEsperado = peso
            End If
        End If

        UltimoPeso = peso
    End Sub
    Sub Mayusculas()
        TextBoxInput.CharacterCasing = CharacterCasing.Upper
    End Sub
    Sub Mesa()
        ' Leer el archivo INI para saber qué mesa buscar
        Dim configPath As String = "config.ini"
        Dim mesaIdArchivo As String = ""
        If File.Exists(configPath) Then
            Dim lines() As String = File.ReadAllLines(configPath)
            For Each line As String In lines
                line = line.Trim()
                If line.StartsWith("Mesa_Id") Then
                    mesaIdArchivo = line.Split("="c)(1).Trim()
                End If
            Next
        End If
        ' Buscar la mesa en la base de datos
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "SELECT Mesas FROM Mesas WHERE IdMesa = @id"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@id", mesaIdArchivo)

                Dim resultado = cmd.ExecuteScalar()
                If resultado IsNot Nothing Then
                    LabelMesa.Text = resultado.ToString()
                Else
                    LabelMesa.Text = "Mesa no encontrada"
                End If
            End Using
        End Using
    End Sub

    ' Variables para validar que ya tenemos todo
    Private Sub TextBoxInput_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxInput.KeyDown
        If e.KeyCode = Keys.Enter Then
            Dim entrada As String = TextBoxInput.Text.Trim()

            ' Validar que la entrada no esté vacía para evitar errores
            If String.IsNullOrEmpty(entrada) Then Exit Sub

            Select Case entrada(0)
                Case "0"c ' Empleado
                    BuscarEmpleado(entrada)
                    CargarRegistros()

                Case "F"c ' Mandril
                    If entrada = ultimoMandril Then
                        ' VALIDACIÓN DE SEGURIDAD
                        If LabelNameTM.Text = "Escanear Numero de Empleado" Or LabelNameTM.Text = "No encontrado" Then
                            LabelAyuda.Text = "⚠️ ERROR: Debe escanear EMPLEADO antes de registrar"
                            LabelAyuda.BackColor = Color.Red
                        Else
                            ' Si hay empleado, entonces sí registramos
                            RegistrarPorCantidad("+" & LabelSP.Text)
                            LabelAyuda.Text = LabelSP.Text & " Piezas registrads ✅"
                            LabelAyuda.BackColor = Color.LawnGreen
                        End If
                    Else
                        ' Si es un mandril nuevo o el primero del día
                        BuscarMandril(entrada)
                        ultimoMandril = entrada
                    End If
                    CargarRegistros()
                Case "+"c ' Registrar piezas personalizadas
                    RegistrarPorCantidad(entrada)
                    CargarRegistros()

                Case Else ' Si no es 0, F o + → puede ser un código de defecto
                    BuscarDefecto(entrada)
                    CargarRegistros()
            End Select

            TextBoxInput.Clear()
            TextBoxInput.Focus()
        End If
    End Sub
    ' --- Validación en BD ---
    Private Function ValidarDatos(mesa As String, tm As String, mandril As String) As Boolean
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()

            ' Validar Mesa
            Dim queryMesa As String = "SELECT COUNT(*) FROM [Mesas] WHERE Mesas = @mesa"
            Using cmdMesa As New SqlCommand(queryMesa, conexion)
                cmdMesa.Parameters.AddWithValue("@mesa", mesa.Trim())
                If Convert.ToInt32(cmdMesa.ExecuteScalar()) = 0 Then
                    LabelAyuda.Text = "La mesa no existe en la base de datos."
                    Return False
                End If
            End Using

            ' Validar TM
            Dim queryTM As String = "SELECT COUNT(*) FROM [User] WHERE [Nombre] = @tm"
            Using cmdTM As New SqlCommand(queryTM, conexion)
                cmdTM.Parameters.AddWithValue("@tm", tm.Trim())
                If Convert.ToInt32(cmdTM.ExecuteScalar()) = 0 Then
                    LabelAyuda.Text = "El trabajador (TM) no existe en la base de datos."
                    Return False
                End If
            End Using

            ' Validar Mandril
            Dim queryMandril As String = "SELECT COUNT(*) FROM [Mandriles] WHERE barcode = @mandril"
            Using cmdMandril As New SqlCommand(queryMandril, conexion)
                cmdMandril.Parameters.AddWithValue("@mandril", mandril.Trim())
                If Convert.ToInt32(cmdMandril.ExecuteScalar()) = 0 Then
                    LabelAyuda.Text = "El mandril no existe en la base de datos."
                    Return False
                End If
            End Using

            ' Si pasa las tres validaciones
            Return True
        End Using
    End Function

    ' --- Registrar piezas por cantidad ---
    Private Sub RegistrarPorCantidad(entrada As String)
        Dim piezasStr As String = entrada.Substring(1) ' quitar el "+"
        Dim piezas As Integer

        If Integer.TryParse(piezasStr, piezas) Then
            ' Validar que haya información suficiente
            If LabelNETM.Text <> "" AndAlso LabelNameTM.Text <> "" AndAlso ultimoMandril <> "" Then
                ' Validar contra la base de datos
                If ValidarDatos(LabelMesa.Text, LabelNameTM.Text, ultimoMandril) Then
                    RegistrarEscaneo(LabelMesa.Text, LabelNameTM.Text, LabelMandril.Text, piezas.ToString())
                    LabelAyuda.Text = "Registro exitoso."
                Else
                    LabelAyuda.Text = "Los datos de Mesa, TM o Mandril no son válidos en la base de datos."
                End If
            Else
                LabelAyuda.Text = "Falta información de empleado o mandril antes de registrar piezas."
            End If
        Else
            LabelAyuda.Text = "Formato inválido en el código: " & entrada
        End If
    End Sub

    Private Sub RegistrarEscaneo(mesa As String, nombreEmpleado As String, mandril As String, cantidadPiezas As String)
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "INSERT INTO RegistrodePiezasEscaneadas 
                               (Fecha, Hora, Mandrel, NDPiezas, Turno, NuMesa, TM) 
                               VALUES (CONVERT(date, GETDATE()), CONVERT(time, GETDATE()), @mandril, @ndpiezas, @turno, @mesa, @tm)"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@mandril", mandril)
                cmd.Parameters.AddWithValue("@ndpiezas", cantidadPiezas)
                cmd.Parameters.AddWithValue("@turno", ObtenerTurno())
                cmd.Parameters.AddWithValue("@mesa", mesa)
                cmd.Parameters.AddWithValue("@tm", nombreEmpleado)
                cmd.ExecuteNonQuery()
            End Using
        End Using
        LabelAyuda.Text = ("Registro realizado correctamente")
    End Sub
    ' -------------------------
    ' Función para buscar empleado
    ' -------------------------
    Private Sub BuscarEmpleado(numeroEmpleado As String)
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "SELECT NumeroDeEmpleado, Nombre FROM [User] WHERE NumeroDeEmpleado = @numero"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@numero", numeroEmpleado)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        LabelNETM.Text = reader("NumeroDeEmpleado").ToString()
                        LabelNameTM.Text = reader("Nombre").ToString()
                    Else
                        LabelNETM.Text = "No encontrado"
                        LabelNameTM.Text = "No encontrado"
                    End If
                End Using
            End Using
        End Using
    End Sub
    ' -------------------------
    ' Función para buscar mandril
    ' -------------------------
    Private Sub BuscarMandril(barcode As String)
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "SELECT Mandril, CantidaddeEmpaque FROM [Mandriles] WHERE Barcode = @barcode AND Area = 'Inspeccion'"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@barcode", barcode)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        LabelMandril.Text = reader("Mandril").ToString()
                        LabelSP.Text = reader("CantidaddeEmpaque").ToString()
                    Else
                        LabelMandril.Text = "No encontrado"
                        LabelSP.Text = "No encontrado"
                    End If
                End Using
            End Using
        End Using
    End Sub
    Private Sub BuscarDefecto(codigoDefecto As String)
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "SELECT Defecto FROM [Defectos] WHERE CodigodeDefecto = @codigo"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@codigo", codigoDefecto)

                Using reader As SqlDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        Dim defectoEncontrado As String = reader("Defecto").ToString()
                        RegistrarDefecto(LabelMesa.Text, LabelNameTM.Text, LabelMandril.Text, codigoDefecto, defectoEncontrado)
                    Else
                        LabelAyuda.Text = ("Código de defecto no encontrado: " & codigoDefecto)
                    End If
                End Using
            End Using
        End Using
    End Sub
    Private Sub RegistrarDefecto(mesa As String, nombreEmpleado As String, mandril As String, codigoDefecto As String, defecto As String)
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "INSERT INTO RegistrodeDefectos 
                               (Fecha, Hora, Mandrel, CodigodeDefecto, Defecto, NuMesa, Turno, TM) 
                               VALUES (CONVERT(date, GETDATE()), CONVERT(time, GETDATE()), @mandril, @codigo, @defecto, @mesa, @turno, @tm)"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@mandril", mandril)
                cmd.Parameters.AddWithValue("@codigo", codigoDefecto)
                cmd.Parameters.AddWithValue("@defecto", defecto)
                cmd.Parameters.AddWithValue("@mesa", mesa)
                cmd.Parameters.AddWithValue("@turno", ObtenerTurno())
                cmd.Parameters.AddWithValue("@tm", nombreEmpleado)
                cmd.ExecuteNonQuery()
            End Using
        End Using
        LabelAyuda.Text = ("Defecto registrado: " & defecto)
    End Sub
    Private Sub CargarRegistros()
        ' Carpeta raíz de la aplicación
        Dim logPath As String = Path.Combine(Application.StartupPath, "CargarRegistros.txt")

        Try
            Using conexion As New SqlConnection(cadenaConexion)
                conexion.Open()
                File.AppendAllText(logPath, $"{DateTime.Now}: Conexión abierta correctamente.{Environment.NewLine}")

                ' --- Piezas buenas agrupadas por Mandrel ---
                Dim queryBuenas As String = "SELECT Mandrel, SUM(TRY_CAST(NDPiezas AS INT)) AS TotalPiezas
                                             FROM RegistrodePiezasEscaneadas
                                             WHERE NuMesa = @mesa 
                                               AND Turno = @turno 
                                               AND TM = @tm
                                               AND CAST(Fecha AS DATE) = CAST(GETDATE() AS DATE)
                                             GROUP BY Mandrel"

                Using cmd As New SqlCommand(queryBuenas, conexion)
                    cmd.Parameters.Add("@mesa", SqlDbType.VarChar).Value = LabelMesa.Text
                    cmd.Parameters.Add("@turno", SqlDbType.VarChar).Value = ObtenerTurno()
                    cmd.Parameters.Add("@tm", SqlDbType.VarChar).Value = LabelNameTM.Text

                    File.AppendAllText(logPath, $"{DateTime.Now}: Ejecutando consulta de piezas buenas.{Environment.NewLine}")

                    Dim adapter As New SqlDataAdapter(cmd)
                    Dim tablaBuenas As New DataTable()
                    adapter.Fill(tablaBuenas)

                    DGVPiezasBuenas.AutoGenerateColumns = True
                    DGVPiezasBuenas.DataSource = tablaBuenas

                    File.AppendAllText(logPath, $"{DateTime.Now}: Consulta completada. Registros cargados: {tablaBuenas.Rows.Count}.{Environment.NewLine}")

                    ' Registrar columnas encontradas
                    For Each col As DataColumn In tablaBuenas.Columns
                        File.AppendAllText(logPath, $"{DateTime.Now}: Columna encontrada -> {col.ColumnName}{Environment.NewLine}")
                    Next
                End Using

                ' --- Defectos ---
                CargarDefectosPivot()
                File.AppendAllText(logPath, $"{DateTime.Now}: Defectos cargados correctamente.{Environment.NewLine}")
            End Using
        Catch ex As Exception
            File.AppendAllText(logPath, $"{DateTime.Now}: ERROR - {ex.Message}{Environment.NewLine}")
            LabelAyuda.Text = ("Error al cargar registros: " & ex.Message)
        End Try
    End Sub
    Private Sub CargarDefectosPivot()
        Using conexion As New SqlConnection(cadenaConexion)
            conexion.Open()
            Dim query As String = "SELECT Defecto, Mandrel
                               FROM RegistrodeDefectos
                               WHERE NuMesa = @mesa AND Turno = @turno AND TM = @tm"
            Using cmd As New SqlCommand(query, conexion)
                cmd.Parameters.AddWithValue("@mesa", LabelMesa.Text)
                cmd.Parameters.AddWithValue("@turno", ObtenerTurno())
                cmd.Parameters.AddWithValue("@tm", LabelNameTM.Text)

                Dim adapter As New SqlDataAdapter(cmd)
                Dim tabla As New DataTable()
                adapter.Fill(tabla)

                ' Crear tabla pivotada
                Dim pivot As New DataTable()
                pivot.Columns.Add("Defecto")

                ' Crear columnas dinámicas por mandril
                Dim mandriles = tabla.AsEnumerable().Select(Function(r) r("Mandrel").ToString()).Distinct().ToList()
                For Each m In mandriles
                    pivot.Columns.Add(m, GetType(Integer))
                Next

                ' Llenar filas por defecto
                Dim defectos = tabla.AsEnumerable().Select(Function(r) r("Defecto").ToString()).Distinct().ToList()
                For Each d In defectos
                    Dim row = pivot.NewRow()
                    row("Defecto") = d
                    For Each m In mandriles
                        Dim count = tabla.AsEnumerable().Count(Function(r) r("Defecto").ToString() = d AndAlso r("Mandrel").ToString() = m)
                        row(m) = count
                    Next
                    pivot.Rows.Add(row)
                Next
                DGVDefectos.DefaultCellStyle.WrapMode = DataGridViewTriState.True
                DGVDefectos.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
                DGVDefectos.DataSource = pivot
            End Using
        End Using
        CalcularTotales()
    End Sub
    Private Sub CalcularTotales()
        Dim totalBuenas As Integer = 0
        Dim totalDefectos As Integer = 0

        ' --- Sumar piezas buenas ---
        For Each row As DataGridViewRow In DGVPiezasBuenas.Rows
            If Not row.IsNewRow AndAlso row.Cells("TotalPiezas").Value IsNot Nothing Then
                totalBuenas += Convert.ToInt32(row.Cells("TotalPiezas").Value)
            End If
        Next

        ' --- Sumar defectos ---
        For Each row As DataGridViewRow In DGVDefectos.Rows
            For Each cell As DataGridViewCell In row.Cells
                ' Ignorar la primera columna (Defecto), solo contar valores numéricos
                If cell.ColumnIndex > 0 AndAlso cell.Value IsNot Nothing Then
                    totalDefectos += Convert.ToInt32(cell.Value)
                End If
            Next
        Next

        ' --- Mostrar resultados en 3 Labels ---
        Dim total As Integer = totalBuenas + totalDefectos
        LabelBuenas.Text = $"Buenas: {totalBuenas}"
        LabelDefectos.Text = $"Defectuosas: {totalDefectos}"
        LabelTotal.Text = $"Total: {total}"
    End Sub
    Private Function ObtenerTurno() As String
        Dim horaActual As Integer = DateTime.Now.Hour
        If horaActual >= 6 AndAlso horaActual < 14 Then
            Return "1"
        ElseIf horaActual >= 14 AndAlso horaActual < 22 Then
            Return "2"
        Else
            Return "3"
        End If
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Intentar reconectar automáticamente cada 10 segundos
        VerificarConexionBascula()
        SetForegroundWindow(Me.Handle)
        TextBoxInput.Focus()
    End Sub

    Private Sub LabelMandril_Click(sender As Object, e As EventArgs) Handles LabelMandril.Click
        InicializarPantalla()
    End Sub
End Class