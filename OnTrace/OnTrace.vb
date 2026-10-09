Imports System.Configuration
Imports System.IO
Imports System.Globalization
Imports System.Text
Imports System.Xml
Imports System.Net
Imports log4net.Repository.Hierarchy
Imports System.ServiceModel
Imports Easy.Logger
Imports log4net
Imports log4net.Appender
Imports log4net.Core
Imports log4net.Layout
Imports log4net.Util

''' <summary>
''' Wrapper class for trace messages to different targets (EventLog, File).
''' </summary>
''' <remarks>
''' All public overloads funnel into a small set of private core methods so that the
''' (large) public surface does not duplicate the routing/formatting logic.
''' </remarks>
Public Class OnTrace

#Region "Declaration"
    Private Shared ReadOnly LogApp As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("AppLoggerSelectLogger"), ConfigurationManager.AppSettings("AppLoggerSelectLogger").ToString(), "AppLogger"))
    Private Shared ReadOnly LogEvent As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("EventLoggerSelectLogger"), ConfigurationManager.AppSettings("EventLoggerSelectLogger").ToString(), "EventLogger"))
    Private Shared ReadOnly LogStartup As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("StartupLoggerSelectLogger"), ConfigurationManager.AppSettings("StartupLoggerSelectLogger").ToString(), "StartupLogger"))

    Private Shared ReadOnly LockObj As New Object
    Private Shared ReadOnly InitLock As New Object

    Private Shared m_blnIsAppLoggerEnabled As Boolean
    Private Shared _bIsLicenseCheckEnabled As Boolean
    Private Shared m_blnIsEventLoggerEnabled As Boolean
    Private Shared m_blnIsStartupLoggerEnabled As Boolean
    Private Shared m_blnConfigured As Boolean
    Private Shared m_blnConfigFailed As Boolean
    Private Shared m_lngTraceLevel As Integer
#End Region

#Region "Properties"
    ''' <summary>
    ''' Triggers whether output is written to "AppLogger".
    ''' </summary>
    Public Shared Property AppLoggerEnabled() As Boolean
        Get
            Return m_blnIsAppLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsAppLoggerEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' If set to false the license file will not be checked.
    ''' </summary>
    Public Shared Property IsLicenseCheckEnabled() As Boolean
        Get
            Return _bIsLicenseCheckEnabled
        End Get
        Set(ByVal value As Boolean)
            _bIsLicenseCheckEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Triggers whether output is written to "EventLogger".
    ''' </summary>
    Public Shared Property EventLoggerEnabled() As Boolean
        Get
            Return m_blnIsEventLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsEventLoggerEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Triggers whether output is written to "StartupLogger".
    ''' </summary>
    Public Shared Property StartupLoggerEnabled() As Boolean
        Get
            Return m_blnIsStartupLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsStartupLoggerEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Level of tracing.
    ''' </summary>
    Public Shared Property TraceLevel() As Integer
        Get
            Return m_lngTraceLevel
        End Get
        Set(ByVal value As Integer)
            m_lngTraceLevel = value
        End Set
    End Property
#End Region

#Region "Enums"
    ''' <summary>
    ''' Levels for showing the program flow.
    ''' </summary>
    ''' <history>
    '''    TD 01.07.07 Creation
    '''    GN 16.01.08 Adding new trace levels 64, 128, 256. Level_All corrected
    ''' </history>
    Public Enum enmTraceLevel
        Level_1 = 1
        Level_2 = 2
        Level_3 = 4
        Level_4 = 8
        Level_5 = 16
        Level_6 = 32
        Level_7 = 64
        Level_8 = 128
        Level_9 = 256
        Level_All = 511
    End Enum

    Private Enum enmTraceType
        TypeError = 1
        TypeWarning = 2
        TypeInformation = 4
        TypeWriteLine = 8
        TypeFatal = 16
    End Enum

    ''' <summary>
    ''' Specifies the trace target to configure (parameter for Reconfigure()).
    ''' </summary>
    Public Enum enmConfigureTarget
        AppLog = 1                  'application log
        EventLog = 2                'event log (eg. WatchDog)
        StartupLog = 4              'startup log (eg. DisplayEngine)
    End Enum

    ''' <summary>
    ''' Specifies the trace target.
    ''' </summary>
    Public Enum enmTarget
        AppLog = 1                  'application log (default)
        EventLog = 2                'event log (eg. WatchDog)
        StartupLog = 4              'startup log (eg. DisplayEngine)
        'combinations of the above
        AppAndEventLog = 3          'application and event log
        AppAndStartupLog = 5        'application and startup log
        EventAndStartupLog = 6      'event and startup log
        AppEventAndStartupLog = 7   'all defined logs
    End Enum
#End Region

#Region "Configuration"

    ''' <summary>
    ''' Raises a warning (in ide/debugging mode only), if the log4net license files are not found.
    ''' The check never throws: a missing license file must not take down the host application.
    ''' </summary>
    Private Shared Sub LicenseCheck()
        ' Only report while running in the IDE (the developer is responsible for linking the
        ' license files to the project).
        If Not System.Diagnostics.Debugger.IsAttached Then Exit Sub
        If Not _bIsLicenseCheckEnabled Then Exit Sub

        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim myOperationContext As OperationContext = OperationContext.Current
        Dim strPath As String = String.Empty
        If httpContext Is Nothing AndAlso myOperationContext Is Nothing Then
            'standalone app
            Try
                strPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
            Catch exNull As NullReferenceException
                System.Diagnostics.Trace.TraceError("OnTrace.LicenseCheck: GetEntryAssembly couldn't be determined")
            End Try
            Try
                If String.IsNullOrEmpty(strPath) Then
                    strPath = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location)
                End If
            Catch exNull As NullReferenceException
                System.Diagnostics.Trace.TraceError("OnTrace.LicenseCheck: neither GetEntryAssembly nor GetCallingAssembly could be determined")
            End Try
        ElseIf httpContext IsNot Nothing Then
            'web app/webservice
            strPath = httpContext.Server.MapPath("~")
        ElseIf myOperationContext IsNot Nothing Then
            'wcf service
            strPath = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath
        End If

        If Not String.IsNullOrEmpty(strPath) Then
            If Not System.IO.File.Exists(System.IO.Path.Combine(strPath, "license_l4n.txt")) Then
                System.Diagnostics.Trace.TraceError("OnTrace.LicenseCheck: log4net license file not found at " & strPath)
            End If
            If Not System.IO.File.Exists(System.IO.Path.Combine(strPath, "notice_l4n.txt")) Then
                System.Diagnostics.Trace.TraceError("OnTrace.LicenseCheck: log4net notice file not found at " & strPath)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Reads the application settings that control logger enablement and the default trace level.
    ''' </summary>
    Private Shared Sub ReadSettingsFromConfig()
        Try
            m_lngTraceLevel = CInt(ConfigurationManager.AppSettings("DefaultTraceLevel"))
        Catch ceex As System.Configuration.ConfigurationErrorsException
            m_lngTraceLevel = 0
        End Try
        Try
            m_blnIsAppLoggerEnabled = CBool(ConfigurationManager.AppSettings("AppLogger"))
        Catch ex As Exception
            m_blnIsAppLoggerEnabled = False
        End Try
        Try
            m_blnIsEventLoggerEnabled = CBool(ConfigurationManager.AppSettings("EventLogger"))
        Catch ex As Exception
            m_blnIsEventLoggerEnabled = False
        End Try
        Try
            m_blnIsStartupLoggerEnabled = CBool(ConfigurationManager.AppSettings("StartupLogger"))
        Catch ex As Exception
            m_blnIsStartupLoggerEnabled = False
        End Try
    End Sub

    ''' <summary>
    ''' Fully (re)configures log4net: license check, XML configuration, appender file-path patching
    ''' and reading the application settings. This is the single initialization routine.
    ''' </summary>
    Private Shared Sub Configure()
        LicenseCheck()

        'reset the configuration, we patch the "File" value below to the freshly configured values
        LogApp.Logger.Repository.ResetConfiguration()
        log4net.Config.XmlConfigurator.Configure()

        PatchAppenderFilePaths()
        ReadSettingsFromConfig()
    End Sub

    ''' <summary>
    ''' Ensures the class has been configured. Thread-safe, lazy and failure-tolerant: if the
    ''' configuration fails, it is logged and the host application is never taken down.
    ''' </summary>
    Private Shared Sub EnsureConfigured()
        If m_blnConfigured OrElse m_blnConfigFailed Then Return
        SyncLock InitLock
            If m_blnConfigured OrElse m_blnConfigFailed Then Return
            Try
                Configure()
                m_blnConfigured = True
            Catch ex As Exception
                m_blnConfigFailed = True
                System.Diagnostics.Trace.TraceError(GetErrorStr(ex, "OnTrace configuration failed; logging may be unavailable"))
            End Try
        End SyncLock
    End Sub

    Private Shared Sub PatchAppenderFilePaths()
        Dim appenders As log4net.Appender.IAppender() = LogApp.Logger.Repository.GetAppenders()
        For Each currentAppender As log4net.Appender.IAppender In appenders
            Try
                If TypeOf currentAppender Is log4net.Appender.FileAppender OrElse
                    TypeOf currentAppender Is log4net.Appender.RollingFileAppender Then
                    Dim fileNameWithPath As String = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName
                    Dim fileName As String = Path.GetFileNameWithoutExtension(fileNameWithPath)
                    Dim folderName As String = New DirectoryInfo(Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory())).Name
                    Dim client As String = GetValueFromAppConfig("Client")
                    If String.IsNullOrEmpty(client) Then
                        client = "Standard"
                    End If

                    Dim baseLogfilePath As String = GetValueFromAppConfig("BasePathLogFile")
                    If String.IsNullOrEmpty(baseLogfilePath) Then
                        baseLogfilePath = "%PROGRAMDATA%\Online Software AG\Logs\"
                    End If
                    If baseLogfilePath.Contains("%") Then
                        Try
                            Dim firstCharIndex As Integer = baseLogfilePath.IndexOf("%")
                            Dim lastCharIndex As Integer = baseLogfilePath.LastIndexOf("%")
                            Dim finalstring As String = baseLogfilePath.Substring(firstCharIndex + 1, (lastCharIndex - 1 - firstCharIndex))
                            Dim envVar As String = Environment.GetEnvironmentVariable(finalstring)
                            If String.IsNullOrEmpty(envVar) Then
                                envVar = "C:\ProgramData\Online Software AG\Logs"
                            End If
                            baseLogfilePath = baseLogfilePath.Replace("%" + finalstring + "%", envVar)
                        Catch ex As Exception
                            baseLogfilePath = "C:\ProgramData\Online Software AG\Logs"
                        End Try
                    End If

                    Dim maybeUnpatchedString As String = CType(currentAppender, log4net.Appender.FileAppender).File
                    If maybeUnpatchedString.Contains(System.AppDomain.CurrentDomain.BaseDirectory()) Then
                        maybeUnpatchedString = maybeUnpatchedString.Replace(System.AppDomain.CurrentDomain.BaseDirectory(), "")
                    End If
                    If maybeUnpatchedString.Contains("{Client}") OrElse maybeUnpatchedString.Contains("{ProgramFolderName}") OrElse maybeUnpatchedString.Contains("{ProgramFilename}") Then
                        If Not maybeUnpatchedString.Contains(baseLogfilePath) Then
                            Dim patchedString As String = maybeUnpatchedString.Replace("{Client}", client).Replace("{ProgramFolderName}", folderName).Replace("{ProgramFilename}", fileName)
                            CType(currentAppender, log4net.Appender.FileAppender).File = baseLogfilePath + patchedString
                            CType(currentAppender, log4net.Appender.FileAppender).ActivateOptions()
                        End If
                    Else
                        If Not maybeUnpatchedString.Contains(baseLogfilePath) Then
                            CType(currentAppender, log4net.Appender.FileAppender).File = baseLogfilePath + maybeUnpatchedString
                            CType(currentAppender, log4net.Appender.FileAppender).ActivateOptions()
                        End If
                    End If
                End If
            Catch ex As Exception
                ' a single un-patchable appender must not abort the whole initialization
                System.Diagnostics.Trace.TraceError(GetErrorStr(ex, "OnTrace: failed to patch appender file path"))
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Changes the log file path and name for AppLog/StartupLog or log name and application name for EventLog,
    ''' and the default trace level (specify "" or -1 to use default setting).
    ''' </summary>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal strFileNameOrLogName As String,
                                       ByVal strFilePathOrApplicationName As String,
                                       ByVal lngTraceLevelDefault As Integer) As Boolean
        EnsureConfigured()
        Try
            Dim blnErrors As Boolean = False

            If strFileNameOrLogName.Trim.Length > 0 OrElse strFilePathOrApplicationName.Trim.Length > 0 Then
                If configureTarget = enmConfigureTarget.EventLog Then
                    Dim appenders As log4net.Appender.IAppender() = LogEvent.Logger.Repository.GetAppenders()
                    Dim astrAppenderOfLogger As String() = GetAppenderNamesForLogger(LogEvent.Logger.Name)
                    Dim intCount As Integer = 0
                    For Each appender As log4net.Appender.IAppender In appenders
                        If Array.IndexOf(astrAppenderOfLogger, appender.Name) <> -1 AndAlso
                            TypeOf appender Is log4net.Appender.EventLogAppender Then
                            If intCount = 0 Then
                                If strFileNameOrLogName.Trim.Length > 0 Then
                                    CType(appender, log4net.Appender.EventLogAppender).LogName = strFileNameOrLogName.Trim
                                End If
                                If strFilePathOrApplicationName.Trim.Length > 0 Then
                                    CType(appender, log4net.Appender.EventLogAppender).ApplicationName = strFilePathOrApplicationName.Trim
                                End If
                                CType(appender, log4net.Appender.EventLogAppender).ActivateOptions()
                                intCount = 1
                            Else
                                blnErrors = True
                                Trace.WriteLine("Reconfigure, error: more than one appender shall be reconfigured to the same eventlog?")
                            End If
                        End If
                    Next
                Else
                    Dim appenders As log4net.Appender.IAppender()
                    Dim astrAppenderOfLogger As String()
                    If configureTarget = enmConfigureTarget.AppLog Then
                        appenders = LogApp.Logger.Repository.GetAppenders()
                        astrAppenderOfLogger = GetAppenderNamesForLogger(LogApp.Logger.Name)
                    Else
                        appenders = LogStartup.Logger.Repository.GetAppenders()
                        astrAppenderOfLogger = GetAppenderNamesForLogger(LogStartup.Logger.Name)
                    End If
                    Dim intCount As Integer = 0
                    For Each appender As log4net.Appender.IAppender In appenders
                        If Array.IndexOf(astrAppenderOfLogger, appender.Name) <> -1 AndAlso
                            (TypeOf appender Is log4net.Appender.FileAppender OrElse
                             TypeOf appender Is log4net.Appender.RollingFileAppender) Then
                            If intCount = 0 Then
                                Dim strFile As String
                                If strFilePathOrApplicationName.Trim.Length > 0 Then
                                    strFile = strFilePathOrApplicationName
                                Else
                                    strFile = Path.GetDirectoryName(CType(appender, log4net.Appender.FileAppender).File)
                                End If
                                If strFileNameOrLogName.Trim.Length > 0 Then
                                    strFile = Path.Combine(strFile, strFileNameOrLogName)
                                Else
                                    strFile = Path.Combine(strFile, Path.GetFileName(CType(appender, log4net.Appender.FileAppender).File))
                                End If
                                CType(appender, log4net.Appender.FileAppender).File = strFile
                                CType(appender, log4net.Appender.FileAppender).ActivateOptions()
                                intCount = 1
                            Else
                                blnErrors = True
                                Trace.WriteLine("Reconfigure, error: more than one appender shall be reconfigured to the same filename?")
                            End If
                        End If
                    Next
                End If
            End If

            If lngTraceLevelDefault <> -1 Then
                m_lngTraceLevel = lngTraceLevelDefault
            End If

            Return Not blnErrors
        Catch ex As Exception
            Trace.TraceError(GetErrorStr(ex, "Reconfigure, ERROR - re-configuring tracing"))
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Changes the log file path and name for AppLog/StartupLog or log name and application name for EventLog.
    ''' </summary>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal strFileNameOrLogName As String,
                                       ByVal strFilePathOrApplicationName As String) As Boolean
        Return Reconfigure(configureTarget, strFileNameOrLogName, strFilePathOrApplicationName, -1)
    End Function

    ''' <summary>
    ''' Changes the default trace level.
    ''' </summary>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal lngTraceLevelDefault As Integer) As Boolean
        Return Reconfigure(configureTarget, "", "", lngTraceLevelDefault)
    End Function

#End Region

#Region "Core"

    ''' <summary>
    ''' Decides whether a message must be written, honoring the trace-level mask.
    ''' Errors and fatal messages are ALWAYS written regardless of the configured trace level.
    ''' </summary>
    Private Shared Function ShouldWrite(ByVal logType As enmTraceType, ByVal lngLevelMessage As enmTraceLevel) As Boolean
        If logType = enmTraceType.TypeError OrElse logType = enmTraceType.TypeFatal Then Return True
        If lngLevelMessage = enmTraceLevel.Level_All Then Return True
        Return (CLng(lngLevelMessage) And CLng(m_lngTraceLevel)) = CLng(lngLevelMessage)
    End Function

    Private Shared Function IsTargetEnabled(ByVal enabled As Boolean, ByVal target As enmTarget, ByVal bit As enmTarget) As Boolean
        Return enabled AndAlso (target And bit) = bit
    End Function

    ''' <summary>
    ''' Resolves the real application call-site (via log4net's stack-boundary mechanism) and stores it
    ''' in the thread context so external layouts can render ClassName/LineNumber/MethodName/FileName.
    ''' </summary>
    Private Shared Sub SetCallerContext(ByVal lngLevelMessage As enmTraceLevel)
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
    End Sub

    Private Shared Sub WriteToLogger(ByVal logger As log4net.ILog, ByVal message As Object, ByVal logType As enmTraceType)
        Select Case logType
            Case enmTraceType.TypeError
                logger.Error(message)
            Case enmTraceType.TypeWarning
                logger.Warn(message)
            Case enmTraceType.TypeInformation
                logger.Info(message)
            Case enmTraceType.TypeFatal
                logger.Fatal(message)
            Case Else
                logger.Debug(message)
        End Select
    End Sub

    ''' <summary>
    ''' Core routing: writes a message to every enabled and selected log target that is affected by
    ''' the configured trace level. Errors and fatal messages bypass the level mask.
    ''' </summary>
    Private Shared Sub WriteToLoggers(ByVal message As Object, ByVal logType As enmTraceType,
                                      ByVal target As enmTarget, ByVal lngLevelMessage As enmTraceLevel)
        EnsureConfigured()
        If Not ShouldWrite(logType, lngLevelMessage) Then Return
        SetCallerContext(lngLevelMessage)

        If IsTargetEnabled(m_blnIsAppLoggerEnabled, target, enmTarget.AppLog) Then
            Try
                WriteToLogger(LogApp, message, logType)
            Catch ex As Exception
                Trace.TraceError(GetErrorStr(ex, "Error writing to AppLogger: "))
            End Try
        End If
        If IsTargetEnabled(m_blnIsEventLoggerEnabled, target, enmTarget.EventLog) Then
            Try
                WriteToLogger(LogEvent, message, logType)
            Catch ex As Exception
                Trace.TraceError(GetErrorStr(ex, "Error writing to EventLogger: "))
            End Try
        End If
        If IsTargetEnabled(m_blnIsStartupLoggerEnabled, target, enmTarget.StartupLog) Then
            Try
                WriteToLogger(LogStartup, message, logType)
            Catch ex As Exception
                Trace.TraceError(GetErrorStr(ex, "Error writing to StartupLogger: "))
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Writes a message to a dynamically created logger that is bound to an external file.
    ''' </summary>
    Private Shared Sub TraceToTarget(ByVal strMessage As Object,
                                     ByVal strFilePath As String,
                                     ByVal strFileName As String,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal logType As enmTraceType)
        strFilePath = GetStrFilePath(strFilePath, strFileName)
        If String.IsNullOrWhiteSpace(strFilePath) Then
            Trace.WriteLine("!!!Warning OnTrace::TraceToTarget Missing path for dynamicLogger '" & strFileName & "', drop message = '" & strMessage & "'")
            Exit Sub
        End If
        If Not ShouldWrite(logType, lngLevelMessage) Then Exit Sub
        SetCallerContext(lngLevelMessage)

        Dim appenderName As String = Path.Combine(strFilePath, strFileName)
        If Not String.IsNullOrWhiteSpace(appenderName) Then
            SyncLock (LockObj)
                Dim appender As IAppender
                Dim dynamicLog As ILog = LogManager.GetLogger(appenderName)
                Dim dynLogAppenders As IAppender = CType(dynamicLog.Logger, Logger).GetAppender(appenderName)
                If dynLogAppenders IsNot Nothing Then
                    appender = dynLogAppenders
                Else
                    appender = GetDynamicAppender(appenderName)
                End If
                Try
                    If CType(dynamicLog.Logger, Logger).Appenders.Contains(appender) = False Then
                        CType(dynamicLog.Logger, Logger).AddAppender(appender)
                    End If
                    CType(dynamicLog.Logger, Logger).Hierarchy.Configured = True
                    WriteToLogger(dynamicLog, strMessage, logType)
                Catch ex As Exception
                    Trace.TraceError(GetErrorStr(ex, "Error writing to Logfile " & appenderName))
                End Try
            End SyncLock
        End If
    End Sub

    ''' <summary>
    ''' We need this method to clean the appenders and loggers after finishing the log process to
    ''' external logfiles.
    ''' </summary>
    Public Shared Sub TraceClear(ByVal strFilePath As String, ByVal strFileName As String)
        strFilePath = GetStrFilePath(strFilePath, strFileName)
        Dim appenderName As String = Path.Combine(strFilePath, strFileName)
        SyncLock (LockObj)
            Dim dynamicLog As log4net.ILog = LogManager.GetLogger(appenderName)
            If dynamicLog IsNot Nothing Then
                If String.IsNullOrWhiteSpace(strFilePath) Then
                    Trace.WriteLine("!!!Warning OnTrace::TraceClear Missing path for dynamicLogger, remove all appender for '" & strFileName & "'!!!")
                End If
                For Each app As IAppender In dynamicLog.Logger.Repository.GetAppenders()
                    'THD is an asyncBufferAppender (logFile as Name) with a rollingFileAppender (timestamp_logFile)
                    If app.Name.EndsWith(strFileName) Then
                        Try
                            CType(dynamicLog.Logger, Logger).RemoveAppender(app)
                        Catch expectedArgExc As ArgumentException
                            'already removed
                        Catch ex As Exception
                            Trace.TraceError(GetErrorStr(ex, "ERROR OnTrace::TraceClear: Remove: " & app.Name))
                        End Try
                        Try
                            app.Close()
                        Catch ex As Exception
                            Trace.TraceError(GetErrorStr(ex, "ERROR OnTrace::TraceClear: Close: " & app.Name))
                        End Try
                    End If
                Next
            End If
        End SyncLock
    End Sub

    Private Shared Function GetDynamicAppender(logFile As String) As IAppender
        Dim lockingType As RollingFileAppender.LockingModelBase = New log4net.Appender.RollingFileAppender.MinimalLock
        Dim layout As PatternLayout = New PatternLayout() With {
                .ConversionPattern = "%date{yyy-MM-dd HH:mm:ss,fff} %-5p [%-2thread] %m%n"
            }
        layout.ActivateOptions()

        Dim timeStamp As String = DateTime.Now.ToString("yyyMMddHHmmssfff")
        Dim appender As RollingFileAppender = New RollingFileAppender() With {
            .Name = timeStamp & "_" & logFile,
            .File = logFile,
            .DatePattern = ".yyyy.MM.dd-HH.log",
            .RollingStyle = log4net.Appender.RollingFileAppender.RollingMode.Size,
            .MaxFileSize = 16,
            .Encoding = System.Text.Encoding.UTF8,
            .MaximumFileSize = "16MB",
            .PreserveLogFileNameExtension = True,
            .StaticLogFileName = True,
            .MaxSizeRollBackups = 24,
            .AppendToFile = True,
            .CountDirection = -1,
            .ImmediateFlush = True,
            .Layout = layout,
            .LockingModel = lockingType
        }
        lockingType.CurrentAppender = appender
        lockingType.ActivateOptions()
        appender.ActivateOptions()

        Dim bufferAppender As AsyncBufferingForwardingAppender = New AsyncBufferingForwardingAppender() With {
            .BufferSize = 1024,
            .Fix = FixFlags.ThreadName Or FixFlags.Message,
            .Lossy = False,
            .Name = logFile
        }
        bufferAppender.AddAppender(appender)
        bufferAppender.ActivateOptions()

        Return bufferAppender
    End Function

#End Region

#Region "TraceError"
    Public Shared Sub TraceErrorFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, args As Object())
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeError, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeError, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object, arg1 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeError, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeError, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeError, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeError, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeError, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeError, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceError(ByVal strMessage As String,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal strFilePath As String,
                                 ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeError)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeError)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeError)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeError)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeError)
    End Sub

    ''' <summary>
    ''' Writes an error message to the AppLog with the given trace level.
    ''' </summary>
    Public Shared Sub TraceError(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(strMessage, enmTraceType.TypeError, enmTarget.AppLog, lngLevelMessage)
    End Sub
    ''' <summary>
    ''' Writes an error message to the AppLog with trace level All.
    ''' </summary>
    Public Shared Sub TraceError(ByVal strMessage As String)
        WriteToLoggers(strMessage, enmTraceType.TypeError, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceError(ByVal strMessage As String,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceError(ByVal strMessage As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeError, enmTarget, lngLevelMessage)
    End Sub
#End Region

#Region "TraceInformation"
    Public Shared Sub TraceInformationFormat(format As String, args As Object())
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeInformation, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeInformation, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeInformation, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeInformation, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeInformation, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeInformation, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeInformation, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeInformation, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub

    ''' <summary>
    ''' Writes an information message to the AppLog with the given trace level.
    ''' </summary>
    Public Shared Sub TraceInformation(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(strMessage, enmTraceType.TypeInformation, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformation(ByVal strMessage As String,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    ''' <summary>
    ''' Writes an information message to the AppLog with trace level All.
    ''' </summary>
    Public Shared Sub TraceInformation(ByVal strMessage As String)
        WriteToLoggers(strMessage, enmTraceType.TypeInformation, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceInformation(ByVal strMessage As String,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceInformation(ByVal strMessage As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeInformation, enmTarget, lngLevelMessage)
    End Sub
#End Region

#Region "TraceWarning"
    Public Shared Sub TraceWarningFormat(format As String, args As Object())
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeWarning, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeWarning, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeWarning, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeWarning, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeWarning, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeWarning, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeWarning, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeWarning, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWarning)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWarning)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWarning)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWarning)
    End Sub

    ''' <summary>
    ''' Writes a warning message to the AppLog with the given trace level.
    ''' </summary>
    Public Shared Sub TraceWarning(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(strMessage, enmTraceType.TypeWarning, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarning(ByVal strMessage As String,
                                   ByVal lngLevelMessage As enmTraceLevel,
                                   ByVal strFilePath As String,
                                   ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWarning)
    End Sub
    ''' <summary>
    ''' Writes a warning message to the AppLog with trace level All.
    ''' </summary>
    Public Shared Sub TraceWarning(ByVal strMessage As String)
        WriteToLoggers(strMessage, enmTraceType.TypeWarning, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceWarning(ByVal strMessage As String,
                                   ByVal lngLevelMessage As enmTraceLevel,
                                   ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceWarning(ByVal strMessage As Object,
                                   ByVal lngLevelMessage As enmTraceLevel,
                                   ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWarning, enmTarget, lngLevelMessage)
    End Sub
#End Region

#Region "TraceFatal"
    Public Shared Sub TraceFatalFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, args As Object())
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeFatal, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeFatal, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object, arg1 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeFatal, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeFatal, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceType.TypeFatal, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceType.TypeFatal, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceType.TypeFatal, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceType.TypeFatal, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatal(ByVal strMessage As String,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal strFilePath As String,
                                 ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeFatal)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, args As Object(),
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeFatal)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeFatal)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeFatal)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal strFilePath As String,
                                       ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeFatal)
    End Sub

    Public Shared Sub TraceFatal(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub TraceFatal(ByVal strMessage As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeFatal, enmTarget, lngLevelMessage)
    End Sub
#End Region

#Region "WriteLine"
    Public Shared Sub WriteLineFormat(format As String, args As Object())
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub
    Public Shared Sub WriteLineFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeInformation)
    End Sub

    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub WriteLine(ByVal strMessage As Object, ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget, enmTraceLevel.Level_All)
    End Sub
    ''' <summary>
    ''' Writes a message to the AppLog.
    ''' </summary>
    Public Shared Sub WriteLine(ByVal strMessage As String)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget.AppLog, enmTraceLevel.Level_All)
    End Sub
    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(strMessage, lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(strMessage, lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub Writeline(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
#End Region

#Region "WriteLineLevel"
    Public Shared Sub WriteLineLevelFormat(format As String, args As Object())
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
    Public Shared Sub WriteLineLevelFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub

    Public Shared Sub WriteLineLevel(ByVal strMessage As String,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal strFilePath As String,
                                     ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
    Public Shared Sub WriteLineLevel(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget.AppLog, lngLevelMessage)
    End Sub
    Public Shared Sub WriteLineLevel(ByVal strMessage As String,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget, lngLevelMessage)
    End Sub
    Public Shared Sub WriteLineLevel(ByVal strMessage As Object,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal enmTarget As enmTarget)
        WriteToLoggers(strMessage, enmTraceType.TypeWriteLine, enmTarget, lngLevelMessage)
    End Sub
#End Region

#Region "Utils"

    Private Shared Function GetStrFilePath(ByVal strFilePath As String, ByRef strFileName As String) As String
        If strFilePath Is Nothing Then
            strFilePath = "C:\LOGFILES"
        End If
        If strFileName Is Nothing Or String.IsNullOrWhiteSpace(strFileName) = True Then
            strFileName = System.IO.Path.GetFileName(strFilePath)
            If strFileName Is Nothing Or String.IsNullOrWhiteSpace(strFileName) = True Then
                strFileName = "Trace.log"
            End If
            strFilePath = System.IO.Path.GetFullPath(strFilePath)
        End If
        Return strFilePath
    End Function

    ''' <summary>
    ''' Gets all appender's names configured for the specified logger.
    ''' </summary>
    Private Shared Function GetAppenderNamesForLogger(ByVal strLoggerName As String) As String()
        Dim astrValues As String() = New String() {}
        Try
            Dim strConfigFilename As String = AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString()
            Dim doc As New System.Xml.XmlDocument
            doc.Load(strConfigFilename)
            Dim nodelist As Xml.XmlNodeList = doc.SelectNodes("configuration/log4net/logger[@name='" & strLoggerName & "']/appender-ref/attribute::ref")
            If Not nodelist Is Nothing Then
                ReDim astrValues(nodelist.Count - 1)
                For i As Integer = 0 To nodelist.Count - 1
                    astrValues(i) = nodelist.Item(i).Value
                Next
            End If
        Catch ex As Exception
            Trace.TraceError(GetErrorStr(ex, "GetAppenderNamesForLogger, error: "))
        End Try
        Return astrValues
    End Function

    ''' <summary>
    ''' Delivers the name of the logfile: path and filename, or the name of the eventlog.
    ''' </summary>
    Public Shared Function GetLogFileName(ByVal enmLogfileType As enmConfigureTarget) As String
        EnsureConfigured()
        Dim strLogfileName As String = String.Empty
        Dim strAppender As String = String.Empty
        Select Case enmLogfileType
            Case enmConfigureTarget.AppLog
                strAppender = "LogFileAppender"
            Case enmConfigureTarget.EventLog
                strAppender = "EventlogAppender"
            Case enmConfigureTarget.StartupLog
                strAppender = "StartupLogFileAppender"
        End Select

        Try
            Dim strConfigFilename As String = AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString()
            Dim doc As New System.Xml.XmlDocument
            doc.Load(strConfigFilename)
            If doc IsNot Nothing Then
                Dim nodelist As Xml.XmlNodeList = doc.SelectNodes("configuration/log4net/appender[@name='" & strAppender & "']/param[@name='File']/attribute::value")
                If Not nodelist Is Nothing Then
                    strLogfileName = nodelist.Item(0).Value
                End If
            End If
        Catch ex As Exception
            Trace.TraceError(GetErrorStr(ex, "GetLogFileName (Configuration for OnTrace is missing?)"))
        End Try
        Return strLogfileName
    End Function

    ''' <summary>
    ''' Applies the configured appender file-path patching and settings. Provided as public API for
    ''' explicit re-initialization; effectively performs the same work as the automatic initialization.
    ''' </summary>
    Public Shared Sub PatchAppenders()
        EnsureConfigured()
    End Sub

    Private Shared Function GetValueFromAppConfig(ByVal key As String) As String
        Dim value As String = ""
        Try
            Dim appSettings As Specialized.NameValueCollection = ConfigurationManager.AppSettings
            value = appSettings(key)
        Catch ex As Exception
            'do nothing here
        End Try
        Return value
    End Function

#End Region

#Region "TraceException"

    Public Shared Sub TraceFatalException(ByVal ex As Exception, filePath As String, fileName As String)
        TraceFatal(GetErrorStr(ex), enmTraceLevel.Level_All, filePath, fileName)
    End Sub
    Public Shared Sub TraceFatalException(ByVal ex As Exception, ByVal description As String, filePath As String, fileName As String)
        TraceFatal(GetErrorStr(ex, description), enmTraceLevel.Level_All, filePath, fileName)
    End Sub
    Public Shared Sub TraceFatalException(ByVal ex As Exception, ByVal description As String)
        TraceFatal(GetErrorStr(ex, description), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub
    Public Shared Sub TraceFatalException(ByVal ex As Exception)
        TraceFatal(GetErrorStr(ex), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub
    Public Shared Sub TraceException(ByVal ex As Exception, filePath As String, fileName As String)
        TraceError(GetErrorStr(ex), enmTraceLevel.Level_All, filePath, fileName)
    End Sub
    Public Shared Sub TraceException(ByVal ex As Exception, ByVal description As String, filePath As String, fileName As String)
        TraceError(GetErrorStr(ex, description), enmTraceLevel.Level_All, filePath, fileName)
    End Sub
    Public Shared Sub TraceException(ByVal ex As Exception, ByVal description As String)
        TraceError(GetErrorStr(ex, description), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub
    Public Shared Sub TraceException(ByVal ex As Exception)
        TraceError(GetErrorStr(ex), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub

#End Region

#Region "SetError Support"

    Public Shared Sub SetError(ByVal ex As Exception, filePath As String, fileName As String)
        TraceError(GetErrorStr(ex), enmTraceLevel.Level_All, filePath, fileName)
    End Sub
    Public Shared Sub SetError(ByVal ex As Exception, ByVal description As String)
        TraceError(GetErrorStr(ex, description), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub
    Public Shared Sub SetError(ByVal ex As Exception)
        TraceError(GetErrorStr(ex), enmTraceLevel.Level_All, enmTarget.AppEventAndStartupLog)
    End Sub
    Public Shared Sub SetError(ByVal ex As Exception, ByVal description As String, filePath As String, fileName As String)
        TraceError(GetErrorStr(ex, description), enmTraceLevel.Level_All, filePath, fileName)
    End Sub

    Public Shared Function GetErrorStr(ByVal ex As Exception) As String
        Return GetErrorStr(ex, "")
    End Function

    Public Shared Function GetErrorStr(ByVal ex As Exception, ByVal description As String) As String
        Dim tempException As Exception = ex
        Dim sb As New StringBuilder
        Const OUTER_SEPERATER As String = "================================================================="
        Const INNER_SEPERATER As String = "-----------------------------------------------------------------"

        sb.Append(description)
        sb.Append(Environment.NewLine)
        sb.Append(OUTER_SEPERATER)
        sb.Append(Environment.NewLine)

        While Not (tempException Is Nothing)
            If tempException.GetType() Is GetType(WebException) Then
                Dim tempWebexception As WebException = DirectCast(tempException, WebException)
                sb.Append(tempException.GetType().ToString())
                sb.Append(": [")
                sb.Append(tempWebexception.Status)
                sb.Append("]")
                sb.Append(tempException.Message)
                sb.Append(Environment.NewLine)
                sb.Append(INNER_SEPERATER)
                sb.Append(Environment.NewLine)
            Else
                sb.Append(tempException.GetType().ToString())
                sb.Append(": ")
                sb.Append(tempException.Message)
                sb.Append(Environment.NewLine)
                sb.Append(INNER_SEPERATER)
                sb.Append(Environment.NewLine)
            End If

            If tempException.Data.Count > 0 Then
                For Each de As DictionaryEntry In tempException.Data
                    sb.AppendFormat("..{0} = {1}", de.Key.ToString(), de.Value)
                Next
                sb.Append(INNER_SEPERATER)
                sb.Append(Environment.NewLine)
            End If

            If String.IsNullOrWhiteSpace(tempException.Source) = False Then
                sb.Append(tempException.Source)
                sb.Append(Environment.NewLine)
                sb.Append(INNER_SEPERATER)
                sb.Append(Environment.NewLine)
            End If
            If String.IsNullOrWhiteSpace(tempException.StackTrace) = False Then
                sb.Append(tempException.StackTrace)
                sb.Append(Environment.NewLine)
                sb.Append(INNER_SEPERATER)
                sb.Append(Environment.NewLine)
            End If
            tempException = tempException.InnerException
        End While
        Return sb.ToString()
    End Function

#End Region

End Class
