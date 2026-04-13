Imports System.Globalization
Imports System.IO.Ports
Imports System.Text.RegularExpressions

Public Class BasculaReader
    Private buffer As String = ""
    Public Event PesoRecibido(peso As Double, crudo As String)

    Private WithEvents puerto As SerialPort

    ' --- Iniciar conexión ---
    Public Function Iniciar() As Boolean
        Try
            If puerto IsNot Nothing AndAlso puerto.IsOpen Then
                Return True
            End If

            ' Validar que el puerto COM3 existe
            If Not SerialPort.GetPortNames().Contains("COM6") Then
                Return False
            End If

            puerto = New SerialPort("COM6", 9600, Parity.None, 7, StopBits.One)
            puerto.Handshake = Handshake.None
            puerto.Encoding = System.Text.Encoding.ASCII
            puerto.NewLine = vbLf
            puerto.ReceivedBytesThreshold = 1
            buffer = ""

            puerto.Open()
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

            ' Procesar solo cuando haya salto de línea
            If buffer.Contains(vbLf) OrElse buffer.Contains(vbCr) Then
                Dim lineas = buffer.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                ' Mantener solo la parte incompleta
                buffer = If(buffer.EndsWith(vbCr) OrElse buffer.EndsWith(vbLf), "", lineas.Last())

                ' Procesar cada línea completa
                For Each l In lineas
                    Dim t = l.Trim()

                    If t.StartsWith("ST") Then
                        Dim match = Regex.Match(t, "([+-]?\d+(\.\d+)?)")
                        If match.Success Then
                            Dim valor As Double = Double.Parse(match.Value, CultureInfo.InvariantCulture)
                            RaiseEvent PesoRecibido(valor, t)
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            ' Ignorar errores
        End Try
    End Sub
End Class