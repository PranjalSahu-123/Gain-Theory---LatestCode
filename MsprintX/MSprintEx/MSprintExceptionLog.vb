Imports System
Imports System.IO

Module MSprintExceptionLog
    Public Function LogMpsrintExException(ByVal exceptionMessage As String, ByVal inputXML As XElement, ByVal outputXML As XElement)
        Try
            Dim PlanLogFile As System.IO.StreamWriter
            PlanLogFile = My.Computer.FileSystem.OpenTextFileWriter(Globals.Ribbons.MSprintExRibbon.ExceptionLogFilePath, True)
            ' PlanLogFile.WriteLine("Here is the first string.")

            ' Using PlanLogFile As StreamWriter = New StreamWriter(Globals.Ribbons.MSprintExRibbon.ExceptionLogFilePath)
            Dim message As String = String.Format("{0} -{1}", Date.Now.ToString(), exceptionMessage)
            PlanLogFile.WriteLine(message)
            Dim inpxml As String = String.Format("Input XML : {0}", inputXML.ToString())
            PlanLogFile.WriteLine(inputXML)
            PlanLogFile.WriteLine(Environment.NewLine)
            Dim opxml As String = String.Format("OutputXML : {0}", outputXML.ToString())
            PlanLogFile.WriteLine(opxml)
            PlanLogFile.Close()
            'End Using
        Catch ex As Exception

        End Try
    End Function
    Public Function LogMpsrintExException(ByVal exceptionMessage As String)
        Try
            Dim PlanLogFile As System.IO.StreamWriter
            PlanLogFile = My.Computer.FileSystem.OpenTextFileWriter(Globals.Ribbons.MSprintExRibbon.ExceptionLogFilePath, True)
            '   Using PlanLogFile As StreamWriter = New StreamWriter(Globals.Ribbons.MSprintExRibbon.ExceptionLogFilePath)
            Dim message As String = String.Format("{0} -{1}", Date.Now.ToString(), exceptionMessage)
            PlanLogFile.WriteLine(message)
            PlanLogFile.Close()
            ' End Using
        Catch ex As Exception

        End Try
    End Function
     
End Module
