Imports System.Globalization
Imports System.IO.Ports
Imports System.Text
Imports System.Text.RegularExpressions

Public Class BasculaReader

    Private buffer As String = ""
    Public Event PesoRecibido(peso As Double, crudo As String)

    Private WithEvents puerto As SerialPort

    ' --- Iniciar conexión ---
    Public Function Iniciar(puertoConfig As String) As Boolean
        Try
            If puerto IsNot Nothing AndAlso puerto.IsOpen Then Return True

            If Not SerialPort.GetPortNames().Contains(puertoConfig) Then
                Return False
            End If

            puerto = New SerialPort(puertoConfig, 9600, Parity.None, 7, StopBits.One)
            puerto.Handshake = Handshake.None
            puerto.Encoding = Encoding.ASCII
            puerto.NewLine = vbLf

            puerto.Open()
            buffer = ""

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    ' --- Detener conexión ---
    Public Sub Detener()
        Try
            If puerto IsNot Nothing AndAlso puerto.IsOpen Then
                puerto.Close()
            End If
        Catch
            ' Ignorar
        End Try
    End Sub

    ' --- Verificar si está conectada ---
    Public Function EstaConectada() As Boolean
        Return puerto IsNot Nothing AndAlso puerto.IsOpen
    End Function

    ' --- Lectura de datos ---
    Private Sub puerto_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles puerto.DataReceived
        Try
            buffer &= puerto.ReadExisting()

            ' Dividir por CR/LF sin eliminar entradas vacías
            Dim partes = buffer.Split({vbCr, vbLf}, StringSplitOptions.None)

            ' Procesar todas menos la última
            For i = 0 To partes.Length - 2
                ProcesarLinea(partes(i).Trim())
            Next

            ' Guardar la última parte (posiblemente incompleta)
            buffer = partes.Last()

        Catch ex As Exception
            ' Ignorar errores
        End Try
    End Sub

    ' --- Procesar una línea completa ---
    Private Sub ProcesarLinea(t As String)
        If String.IsNullOrWhiteSpace(t) Then Exit Sub

        If t.StartsWith("ST") Then
            Dim match = Regex.Match(t, "([+-]?\d+(\.\d+)?)")
            If match.Success Then
                Dim valor As Double = Double.Parse(match.Value, CultureInfo.InvariantCulture)
                RaiseEvent PesoRecibido(valor, t)
            End If
        End If
    End Sub

End Class