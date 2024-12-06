Public Class CustomLogUtil
    Public logOutputDir As String = String.Empty
    Public Sub New(ByVal logDir As String)
        logDir = logOutputDir
    End Sub
#Region "CustomLog"
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="logTitle"></param>
    ''' <param name="logMsg"></param>
    ''' <param name="logStrackstrace"></param>
    ''' 
    Public Shared Sub Heading(logTitle, Optional logMsg = Nothing, Optional logStrackstrace = Nothing)
        'Flow
        '(1)outputDirectory
        '(2)NO of Arguments
        '(3)NewLine
        '(4)Date and Time of log-msg
        '(5)Type and Title of Error
        '(6)Error Message
        '(7)Error StackTrace
        Dim header As String = "============"
        Dim footer As String = "============"
        'Dim LogDir As MonarchLog
        Dim LogDir As String = $"{Config.configObj.LogOutputDirectory}\BEC_Automation_{Date.Now.Day.ToString + "-" + Date.Now.Month.ToString + "-" + Date.Now.Year.ToString}.txt"
        Dim Type As String
        'logMsg.ToString()
        'logStrackstrace.ToString()

        Type = "INFO"

        IO.File.AppendAllText(
            LogDir,
            String.Format(
            "{0}[{1}][{2}]{3}{4}[{5}][{6}]{7}{8}[{9}][{10}]{11}",
            System.Environment.NewLine,
            DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"),
            Type, " " + header + " ",
            System.Environment.NewLine,
            DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"),
            Type, " " + logTitle + " ",
            System.Environment.NewLine,
            DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"),
            Type, " " + footer + " "
))


        'Process.Start(LogDir)
    End Sub

    Public Shared Sub Log(logTitle, Optional logMsg = Nothing, Optional logStrackstrace = Nothing)
        'Flow
        '(1)outputDirectory
        '(2)NO of Arguments
        '(3)NewLine
        '(4)Date and Time of log-msg
        '(5)Type and Title of Error
        '(6)Error Message
        '(7)Error StackTrace

        'Dim LogDir As MonarchLog

        '15th Nov 2024
        If Config.configObj Is Nothing Then
            Exit Sub
        End If

        Dim LogDir As String = $"{Config.configObj.LogOutputDirectory}\BEC_Automation_{Date.Now.Day.ToString + "-" + Date.Now.Month.ToString + "-" + Date.Now.Year.ToString}.txt"
        Dim Type As String
        'logMsg.ToString()
        'logStrackstrace.ToString()
#Region "Type_Validation"
        If logMsg IsNot Nothing And logStrackstrace IsNot Nothing Then
            Type = "ERROR"
        Else
            Type = "INFO"
        End If

#End Region
        If Type = "ERROR" Then
            IO.File.AppendAllText(
            LogDir,'(1)
            String.Format(
            "{0}[{1}][{2}]{3}{4}{5}",'(2)
            System.Environment.NewLine,'(3)
            DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"),'(4)
            Type, " " + logTitle + " | ",'(5) 
            logMsg + " | ",'(6) 
            logStrackstrace + " | "'(7) 
            ))
        Else
            IO.File.AppendAllText(
            LogDir,'(1)
            String.Format(
            "{0}[{1}][{2}]{3}",'(2)
            System.Environment.NewLine,'(3)
            DateTime.Now.ToString("MM-dd-yyyy hh:mm:ss"),'(4)
            Type, " " + logTitle + " | "'(5) 
             ))

        End If





        'Process.Start(LogDir)
    End Sub
#End Region

End Class
