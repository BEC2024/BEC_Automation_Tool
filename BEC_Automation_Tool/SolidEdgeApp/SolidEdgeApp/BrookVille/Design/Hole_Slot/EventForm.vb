Imports System
Imports System.Runtime.InteropServices

Public Class EventForm

    Private _application As SolidEdgeFramework.Application
    Private _applicationEvents As SolidEdgeFramework.ISEApplicationEvents_Event

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub EventForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        _application = DirectCast(Marshal.GetActiveObject("SolidEdge.Application"), SolidEdgeFramework.Application)

        ' Connect to application events.
        _applicationEvents = CType(_application.ApplicationEvents, SolidEdgeFramework.ISEApplicationEvents_Event)
        AddHandler _applicationEvents.AfterDocumentSave, AddressOf _applicationEvents_AfterDocumentSave
        AddHandler _applicationEvents.BeforeDocumentSave, AddressOf _applicationEvents_BeforeDocumentSave
        AddHandler _applicationEvents.BeforeCommandRun, AddressOf _applicationEvents_BeforeCommandRun

    End Sub

    Private Sub _applicationEvents_AfterDocumentSave(ByVal theDocument As Object)
        ' Handle AfterDocumentSave.
    End Sub

    Private Sub _applicationEvents_BeforeCommandRun(ByVal theCommandID As Integer)
        ' Handle Before command run.

        If theCommandID = 25011 Or theCommandID = 25012 Or theCommandID = 25013 Or theCommandID = 25014 Or theCommandID = 25015 Or theCommandID = 25016 Or theCommandID = 25018 Then
            'Debug.Print(" circle Command center executed")
            'MsgBox("Before circle command executed")
            'Dim tmr As New System.Timers.Timer

            Dim oMessageBoxForm As MessageBoxForm = New MessageBoxForm($"Information", "USE HOLE FEATURE TO CREATE HOLE", MessageBoxForm.MessageType.InformationMessage, True)

            Try
                oMessageBoxForm.Show()
            Catch ex As Exception

            End Try


            'MessageBox.Show("USE HOLE FEATURE TO CREATE HOLE.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End If
    End Sub


    Private Sub _applicationEvents_BeforeDocumentSave(ByVal theDocument As Object)
        ' Handle BeforeDocumentSave.
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs)
        ' Disconnect to application events.
        AddHandler _applicationEvents.BeforeDocumentSave, AddressOf _applicationEvents_BeforeDocumentSave
        AddHandler _applicationEvents.AfterDocumentSave, AddressOf _applicationEvents_AfterDocumentSave

        _applicationEvents = Nothing
    End Sub

End Class