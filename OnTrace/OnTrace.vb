Imports System.Configuration
Imports System.IO
Imports System.Globalization
Imports log4net.Repository.Hierarchy
Imports System.Net
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
''' <remarks></remarks>
''' <history>
'''   TD 03.07.07 created
'''   TD 23.10.07 replaced optional parameters with overloaded methods
'''   GN 16.01.08 Adding new trace levels 64, 128, 256. Level_All corrected
'''   MM 21.07.10 Adding writeline ctors
'''   MA 21.02.18 Added ability to select the logger for an default logger, so the logger name changed and you get a better separation of logs if you want to use multiple udp appenders
''' </history>
Public Class OnTrace

#Region "Declaration"
    Private Shared ReadOnly LogApp As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("AppLoggerSelectLogger"), ConfigurationManager.AppSettings("AppLoggerSelectLogger").ToString(), "AppLogger"))
    Private Shared ReadOnly LogEvent As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("EventLoggerSelectLogger"), ConfigurationManager.AppSettings("EventLoggerSelectLogger").ToString(), "EventLogger"))
    Private Shared ReadOnly LogStartup As log4net.ILog = log4net.LogManager.GetLogger(If(ConfigurationManager.AppSettings.AllKeys.Contains("StartupLoggerSelectLogger"), ConfigurationManager.AppSettings("StartupLoggerSelectLogger").ToString(), "StartupLogger"))

    Private Shared ReadOnly LockObj As New Object
    Private Shared m_blnIsAppLoggerEnabled As Boolean
    Private Shared _bIsLicenseCheckEnabled As Boolean
    Private Shared m_blnIsEventLoggerEnabled As Boolean
    Private Shared m_blnIsStartupLoggerEnabled As Boolean
    Private Shared m_blnConfigured As Boolean

#End Region

#Region "Properties"
    Private Shared m_lngTraceLevel As Integer
    ''' <summary>
    ''' Triggers whether output is written to "AppLogger"
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property AppLoggerEnabled() As Boolean
        Get
            Return m_blnIsAppLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsAppLoggerEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' if set to false the license file will not be checked
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property IsLicenseCheckEnabled() As Boolean
        Get
            Return _bIsLicenseCheckEnabled
        End Get
        Set(ByVal value As Boolean)
            _bIsLicenseCheckEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Triggers whether output is written to "EventLogger"
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property EventLoggerEnabled() As Boolean
        Get
            Return m_blnIsEventLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsEventLoggerEnabled = value
        End Set
    End Property

    ''' <summary>
    ''' Triggers whether output is written to "StartupLogger"
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Property StartupLoggerEnabled() As Boolean
        Get
            Return m_blnIsStartupLoggerEnabled
        End Get
        Set(ByVal value As Boolean)
            m_blnIsStartupLoggerEnabled = value
        End Set
    End Property

    'necessary? or reading out of application config?
    ''' <summary>
    '''     Level of tracing
    ''' </summary>
    ''' <value>
    '''     <para>
    '''         Level to set
    '''     </para>
    ''' </value>
    ''' <history>
    '''     GN 03.11.05 Introducing a trace level for Traces
    ''' </history>
    Public Shared Property TraceLevel() As Integer
        Get
            TraceLevel = m_lngTraceLevel
        End Get
        Set(ByVal value As Integer)
            m_lngTraceLevel = value
        End Set
    End Property
#End Region

#Region "Enums"
    ''' <summary>
    ''' Levels for showing the program flow
    ''' </summary>
    ''' <remarks>
    ''' The usage of the levels is different in different applications. The following lists describe 
    ''' the levels used in these applications.
    ''' DisplayEngine.Net:
    '''     Level_1     1 = Positioning information about forms in general
    '''     Level_2     2 = Information about the ticker
    '''     Level_3     3 = NOT USED (formerly used to display the playlist contents)
    '''     Level_4     8 = Timer Scheduler, Cyclic 
    '''     Level_5    16 = Positioning and Resizing information of forms, ResizeWindow(), Cyclic in CheckWindowPositionAndSize
    '''     Level_6    32 = Timer FileChecker, cyclic
    '''     Level_7    64 = not used yet
    ''' DisplayEngineConnector:
    ''' 
    ''' DisplayEngineEventhandler:
    ''' </remarks>
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
    ''' <remarks>
    ''' same basic values as in enmTarget
    ''' </remarks>
    Public Enum enmConfigureTarget
        AppLog = 1                  'application log
        EventLog = 2                'event log (eg. WatchDog)
        StartupLog = 4              'startup log (eg. DisplayEngine)
    End Enum

    ''' <summary>
    ''' Specifies the trace target.
    ''' </summary>
    ''' <remarks></remarks>
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
    ''' static constructor
    ''' </summary>
    Shared Sub New()
        _bIsLicenseCheckEnabled = True

        PatchAppenders()
    End Sub
    ''' <summary>
    ''' Raises an exception (in ide/debugging mode only), if the log4net license files are not found
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub LicenseCheck()

        'we throw only exceptions while running in ide 
        '(developer is responsible for linking the license files to the projects)
        If Not System.Diagnostics.Debugger.IsAttached Then Exit Sub
        ' if License should not be checked: calling web services via IISEXPRESS leads to problems
        If Not _bIsLicenseCheckEnabled Then Exit Sub
        'running as web or standalone application?
        Dim httpContext As System.Web.HttpContext = System.Web.HttpContext.Current
        Dim myOperationContext As System.ServiceModel.OperationContext = OperationContext.Current
        Dim strPath As String = String.Empty
        If httpContext Is Nothing AndAlso myOperationContext Is Nothing Then
            'GN 02.10.13
            'In case of a testrun, the EntryAssembly isn't defined
            'standalone app
            Try
                strPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)
            Catch exNull As NullReferenceException
                Debug.WriteLine("GetEntryAssembly couldn't be determined")
            End Try
            Try
                'In this case, we should have the Try for CallingAssembly
                If String.IsNullOrEmpty(strPath) Then
                    Debug.WriteLine("Trying with GetCallingAssembly")
                    strPath = Path.GetDirectoryName(System.Reflection.Assembly.GetCallingAssembly().Location)
                End If
            Catch exNull As NullReferenceException
                Debug.WriteLine("FATAL - Neither GetEntryAssembly nor GetCallingAssembly could be determined")
            End Try

        ElseIf httpContext IsNot Nothing Then
            'web app/webservice
            strPath = httpContext.Server.MapPath("~")

        ElseIf myOperationContext IsNot Nothing Then
            'wcf service
            strPath = System.Web.Hosting.HostingEnvironment.ApplicationPhysicalPath
        End If
        Debug.WriteLine("searching license file from: " + strPath)
        If Not System.IO.File.Exists(System.IO.Path.Combine(strPath, "license_l4n.txt")) Then Throw New FileNotFoundException("Could not find log4net license file", "license_l4n.txt")
        If Not System.IO.File.Exists(System.IO.Path.Combine(strPath, "notice_l4n.txt")) Then Throw New FileNotFoundException("Could not find log4net notice file", "notice_l4n.txt")

    End Sub

    ''' <summary>
    ''' This method must be called from all public methods to ensure proper intialization
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Sub Configure()

        'check for existence of license files
        LicenseCheck()

        'configure all loggers
        log4net.Config.XmlConfigurator.Configure()

        'read default trace level
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

        'mark class as configured
        m_blnConfigured = True

    End Sub
    ''' <summary>
    ''' Changes the log file path and name for AppLog/StartupLog or log name and application name for EventLog,
    ''' and the default trace level (specify "" or -1 to use default setting)
    ''' </summary>
    ''' <param name="configureTarget">target to configure</param>
    ''' <param name="strFileNameOrLogName">filename (AppLog/StartupLog) or logname (EventLog)</param>
    ''' <param name="strFilePathOrApplicationName">path (AppLog/StartupLog) or application name (EventLog)</param>
    ''' <param name="lngTraceLevelDefault">new default trace level</param>
    ''' <returns>success / error</returns>
    ''' <remarks></remarks>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal strFileNameOrLogName As String,
                                       ByVal strFilePathOrApplicationName As String,
                                       ByVal lngTraceLevelDefault As Integer) As Boolean

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        Try
            Dim blnErrors As Boolean = False

            'at least one parameter must be set to reconfigure an appender
            If strFileNameOrLogName.Trim.Length > 0 OrElse strFilePathOrApplicationName.Trim.Length > 0 Then

                If configureTarget = enmConfigureTarget.EventLog Then

                    Dim appenders As log4net.Appender.IAppender()
                    appenders = LogEvent.Logger.Repository.GetAppenders()
                    Dim astrAppenderOfLogger As String()
                    astrAppenderOfLogger = GetAppenderNamesForLogger(LogEvent.Logger.Name)
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
                                'there should be only one appender configured
                                intCount = 1
                            Else
                                blnErrors = True
                                Trace.WriteLine("Reconfigure, error: more than one appender shall be reconfigured to the same eventlog?")
                            End If
                        End If
                    Next

                Else
                    'reconfigure either AppLog or StartupLog

                    'at least one change specified
                    Dim appenders As log4net.Appender.IAppender()
                    Dim astrAppenderOfLogger As String()
                    If configureTarget = enmConfigureTarget.AppLog Then
                        'reconfigure app log
                        appenders = LogApp.Logger.Repository.GetAppenders()
                        astrAppenderOfLogger = GetAppenderNamesForLogger(LogApp.Logger.Name)
                    Else
                        'reconfigure startup log
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
                                'determine directory
                                If strFilePathOrApplicationName.Trim.Length > 0 Then
                                    strFile = strFilePathOrApplicationName
                                Else
                                    strFile = Path.GetDirectoryName(CType(appender, log4net.Appender.FileAppender).File)
                                End If
                                'append file name
                                If strFileNameOrLogName.Trim.Length > 0 Then
                                    strFile = Path.Combine(strFile, strFileNameOrLogName)
                                Else
                                    strFile = Path.Combine(strFile, Path.GetFileName(CType(appender, log4net.Appender.FileAppender).File))
                                End If
                                CType(appender, log4net.Appender.FileAppender).File = strFile
                                CType(appender, log4net.Appender.FileAppender).ActivateOptions()
                                'there should be only one appender configured
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
            'We can't do anything else
            Trace.WriteLine("Reconfigure, ERROR - re-configuring tracing: " & GetErrorStr(ex))
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Changes the log file path and name for AppLog/StartupLog or log name and application name for EventLog 
    ''' (using the default trace level) (specify "" or -1 to use default setting)
    ''' </summary>
    ''' <param name="configureTarget">target to configure</param>
    ''' <param name="strFileNameOrLogName">filename (AppLog/StartupLog) or logname (EventLog)</param>
    ''' <param name="strFilePathOrApplicationName">path (AppLog/StartupLog) or application name (EventLog)</param>
    ''' <returns>success / error</returns>
    ''' <remarks></remarks>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal strFileNameOrLogName As String,
                                       ByVal strFilePathOrApplicationName As String) As Boolean
        Reconfigure(configureTarget, strFileNameOrLogName, strFilePathOrApplicationName, -1)
    End Function

    ''' <summary>
    ''' Changes the default trace level (using filename (AppLog/StartupLog) or logname (EventLog), 
    ''' and path (AppLog/StartupLog) or application name (EventLog))
    ''' </summary>
    ''' <param name="configureTarget">target to configure</param>
    ''' <param name="lngTraceLevelDefault">new default trace level</param>
    ''' <returns>success / error</returns>
    ''' <remarks></remarks>
    Public Shared Function Reconfigure(ByVal configureTarget As enmConfigureTarget,
                                       ByVal lngTraceLevelDefault As Integer) As Boolean
        Reconfigure(configureTarget, "", "", lngTraceLevelDefault)
    End Function

#End Region

#Region "Dynamic Logger Support"
    ''' <summary>
    ''' We need this Method to clen the appenders and loggers after finishing the log process to external logfiles
    ''' </summary>
    ''' <param name="strFilePath"></param>
    ''' <param name="strFileName"></param>
    ''' <remarks></remarks>
    ''' <history>
    ''' 11.10.2012 NB Created
    '''</history>
    Public Shared Sub TraceClear(ByVal strFilePath As String,
                     ByVal strFileName As String)

        'build the logger and appender name from the given paramaters
        strFilePath = GetStrFilePath(strFilePath, strFileName)
        Dim appenderName As String = Path.Combine(strFilePath, strFileName)
        'if logger does not exists it will be created else it will get the logger
        SyncLock (LockObj)
            Dim dynamicLog As log4net.ILog
            'if logger does not exists it will be created else it will get the logger
            dynamicLog = LogManager.GetLogger(appenderName)
            If dynamicLog IsNot Nothing Then
                Dim dynLogAppender As IAppender = CType(dynamicLog.Logger, Logger).GetAppender(appenderName)
                If dynLogAppender IsNot Nothing Then
                    CType(dynamicLog.Logger, Logger).RemoveAppender(dynLogAppender)
                    dynLogAppender.Close()
                End If
            End If
        End SyncLock
    End Sub
    Private Shared Sub TraceToTarget(ByVal strMessage As String,
                                        ByVal strFilePath As String,
                                        ByVal strFileName As String,
                                        ByVal lngLevelMessage As enmTraceLevel,
                                        ByVal logType As enmTraceType
                                        )
        strFilePath = GetStrFilePath(strFilePath, strFileName)
        If Not m_blnConfigured Then Configure()
        'check loglevel
        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
        Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        'Dim processid As String = "[" & Process.GetCurrentProcess().Id.ToString("00000000") & "]"
        'Dim appenderName As String = Path.Combine(strFilePath, processid, strFileName)
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        Dim appenderName As String = Path.Combine(strFilePath, strFileName)
        If String.IsNullOrWhiteSpace(appenderName) = False Then
            SyncLock (LockObj)
                Dim appender As IAppender
                Dim dynamicLog As ILog
                'if logger does not exists it will be created else it will get the logger
                dynamicLog = LogManager.GetLogger(appenderName)
                Dim dynLogAppenders As IAppender = CType(dynamicLog.Logger, Logger).GetAppender(appenderName)
                If dynLogAppenders IsNot Nothing Then
                    appender = dynLogAppenders
                Else
                    appender = GetDynamicAppender(appenderName)
                End If

                Try
                    'add appender to hierarchy
                    If CType(dynamicLog.Logger, Logger).Appenders.Contains(appender) = False Then
                        CType(dynamicLog.Logger, Logger).AddAppender(appender)
                    End If
                    CType(dynamicLog.Logger, Logger).Hierarchy.Configured = True
                    'Write to log
                    Select Case logType
                        Case enmTraceType.TypeError
                            dynamicLog.Error(strMessage)
                        Case enmTraceType.TypeInformation
                            dynamicLog.Info(strMessage)
                        Case enmTraceType.TypeWarning
                            dynamicLog.Warn(strMessage)
                        Case enmTraceType.TypeWriteLine
                            dynamicLog.Debug(strMessage)
                        Case enmTraceType.TypeFatal
                            dynamicLog.Fatal(strMessage)
                        Case Else
                            dynamicLog.Debug(strMessage)
                    End Select
                Catch ex As Exception
                    Trace.WriteLine("Error writing to Logfile " & appenderName & ": " & GetErrorStr(ex))
                End Try
            End SyncLock
        End If
    End Sub
    ''' <summary>
    ''' This method writes a log file to a defined path
    ''' </summary>
    ''' <remarks>
    '''</remarks>
    ''' <history>
    ''' NB: 16.2.2012 Created
    '''</history>
    Private Shared Sub TraceToTarget(ByVal strMessage As Object,
                                        ByVal strFilePath As String,
                                        ByVal strFileName As String,
                                        ByVal lngLevelMessage As enmTraceLevel,
                                        ByVal logType As enmTraceType
                                        )
        strFilePath = GetStrFilePath(strFilePath, strFileName)
        If Not m_blnConfigured Then Configure()
        'check loglevel
        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
        Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        'Dim processid As String = "[" & Process.GetCurrentProcess().Id.ToString("00000000") & "]"
        'Dim appenderName As String = Path.Combine(strFilePath, processid, strFileName)
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        Dim appenderName As String = Path.Combine(strFilePath, strFileName)
        If String.IsNullOrWhiteSpace(appenderName) = False Then
            SyncLock (LockObj)
                Dim appender As IAppender
                Dim dynamicLog As ILog
                'if logger does not exists it will be created else it will get the logger
                dynamicLog = LogManager.GetLogger(appenderName)
                Dim dynLogAppenders As IAppender = CType(dynamicLog.Logger, Logger).GetAppender(appenderName)
                If dynLogAppenders IsNot Nothing Then
                    appender = dynLogAppenders
                Else
                    appender = GetDynamicAppender(appenderName)
                End If

                Try
                    'add appender to hierarchy
                    If CType(dynamicLog.Logger, Logger).Appenders.Contains(appender) = False Then
                        CType(dynamicLog.Logger, Logger).AddAppender(appender)
                    End If
                    CType(dynamicLog.Logger, Logger).Hierarchy.Configured = True
                    'Write to log
                    Select Case logType
                        Case enmTraceType.TypeError
                            dynamicLog.Error(strMessage)
                        Case enmTraceType.TypeInformation
                            dynamicLog.Info(strMessage)
                        Case enmTraceType.TypeWarning
                            dynamicLog.Warn(strMessage)
                        Case enmTraceType.TypeWriteLine
                            dynamicLog.Debug(strMessage)
                        Case enmTraceType.TypeFatal
                            dynamicLog.Fatal(strMessage)
                        Case Else
                            dynamicLog.Debug(strMessage)
                    End Select
                Catch ex As Exception
                    Trace.WriteLine("Error writing to Logfile " & appenderName & ": " & GetErrorStr(ex))
                End Try
            End SyncLock
        End If
    End Sub
    Private Shared Function GetDynamicAppender(logFile As String) As IAppender
        'filter
        'Dim filter As New log4net.Filter.LevelRangeFilter With {
        '        .LevelMax = log4net.Core.Level.Warn,
        '        .LevelMin = log4net.Core.Level.Debug
        '    }

        Dim lockingType As RollingFileAppender.LockingModelBase = New log4net.Appender.RollingFileAppender.MinimalLock
        Dim layout As PatternLayout = New PatternLayout() With {
                .ConversionPattern = "%date{yyy-MM-dd HH:mm:ss,fff} %-5p [%-2thread] %m%n"
            }
        layout.ActivateOptions()

        Dim timeStamp As String = DateTime.Now.ToString("yyyMMddHHmmssfff")
        Dim appender As RollingFileAppender = New RollingFileAppender() With {
            .Name = timeStamp,
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
        'appender.AddFilter(filter)
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
#Region "TraceError Format Support"
    Public Shared Sub TraceErrorFormat(ByVal format As String, args As Object(),
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, args As Object())
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object, arg1 As Object)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, args As Object(),
                             ByVal lngLevelMessage As enmTraceLevel)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceErrorFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceError(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
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

#End Region
    ''' <summary>
    ''' Writes an error message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceError(ByVal strMessage As String,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceError(strMessage, lngLevelMessage, enmTarget.AppLog)
    End Sub
    ''' <summary>
    ''' Writes an error message to the AppLog with trace level All.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceError(ByVal strMessage As String)
        TraceError(strMessage, enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceError(ByVal strMessage As String,
                             ByVal lngLevelMessage As enmTraceLevel,
                             ByVal enmTarget As enmTarget)
        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Writes an error message to the trace targets.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <param name="enmTarget">trace targets</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceError(ByVal strMessage As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub

        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Error(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If

    End Sub


#End Region

#Region "TraceInformation"
#Region "TraceInformation Format Support"
    Public Shared Sub TraceInformationFormat(format As String, args As Object())
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceInformationFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceInformation(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget)
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

#End Region
    ''' <summary>
    ''' Writes an information message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceInformation(ByVal strMessage As String,
                                       ByVal lngLevelMessage As enmTraceLevel)
        TraceInformation(strMessage, lngLevelMessage, enmTarget.AppLog)
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
    ''' <param name="strMessage">message to write</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceInformation(ByVal strMessage As String)
        TraceInformation(strMessage, enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceInformation(ByVal strMessage As String,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal enmTarget As enmTarget)
        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Writes an information message to the trace targets.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <param name="enmTarget">trace targets</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceInformation(ByVal strMessage As Object,
                                       ByVal lngLevelMessage As enmTraceLevel,
                                       ByVal enmTarget As enmTarget)

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Info(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If

    End Sub

#End Region

#Region "TraceWarning"
#Region "Trace Warning Format Support"
    Public Shared Sub TraceWarningFormat(format As String, args As Object())
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, args As Object(), ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub TraceWarningFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        TraceWarning(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget)
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
#End Region
    ''' <summary>
    ''' Writes a warning message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceWarning(ByVal strMessage As String,
                                   ByVal lngLevelMessage As enmTraceLevel)
        TraceWarning(strMessage, lngLevelMessage, enmTarget.AppLog)
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
    ''' <param name="strMessage">message to write</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceWarning(ByVal strMessage As String)
        TraceWarning(strMessage, enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceWarning(ByVal strMessage As String,
                                  ByVal lngLevelMessage As enmTraceLevel,
                                  ByVal enmTarget As enmTarget)

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Writes a warning message to the trace targets.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level</param>
    ''' <param name="enmTarget">trace targets</param>
    ''' <remarks></remarks>
    Public Shared Sub TraceWarning(ByVal strMessage As Object,
                                   ByVal lngLevelMessage As enmTraceLevel,
                                   ByVal enmTarget As enmTarget)
        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                'write to application log
                LogApp.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                'write to eventlog
                LogEvent.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                'write to startup log
                LogStartup.Warn(strMessage)
            Catch ex As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
            End Try
        End If

    End Sub
#End Region

#Region "TraceFatal Support"
#Region "TraceError Format Support"
    Public Shared Sub TraceFatalFormat(ByVal format As String, args As Object(),
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel,
                                 ByVal enmTarget As enmTarget)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, args As Object())
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object, arg1 As Object)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(format As String, arg0 As Object, arg1 As Object, arg2 As Object)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), enmTraceLevel.Level_All, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, args As Object(),
                             ByVal lngLevelMessage As enmTraceLevel)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, args), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1), lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub TraceFatalFormat(ByVal format As String, arg0 As Object, arg1 As Object, arg2 As Object,
                                 ByVal lngLevelMessage As enmTraceLevel)
        TraceFatal(New SystemStringFormat(CultureInfo.InvariantCulture, format, arg0, arg1, arg2), lngLevelMessage, enmTarget.AppLog)
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

#End Region
    Public Shared Sub TraceFatal(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
           Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                LogApp.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.TraceError("Error writing to AppLogger: " & GetErrorStr(ex1))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                LogEvent.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex1))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                LogStartup.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex1))
            End Try
        End If
    End Sub
    Public Shared Sub TraceFatal(ByVal strMessage As Object, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        If Not m_blnConfigured Then Configure()

        If Not lngLevelMessage = enmTraceLevel.Level_All AndAlso
          Not ((lngLevelMessage And m_lngTraceLevel) = lngLevelMessage) Then Exit Sub
        Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
        log4net.ThreadContext.Properties("LogLevel") = lngLevelMessage
        log4net.ThreadContext.Properties("ClassName") = info.ClassName
        log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
        log4net.ThreadContext.Properties("MethodName") = info.MethodName
        log4net.ThreadContext.Properties("FileName") = info.FileName
        If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
            Try
                LogApp.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.TraceError("Error writing to AppLogger: " & GetErrorStr(ex1))
            End Try
        End If
        If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
            Try
                LogEvent.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex1))
            End Try
        End If
        If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
            Try
                LogStartup.Fatal(strMessage)
            Catch ex1 As Exception
                Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex1))
            End Try
        End If
    End Sub
#End Region

#Region "WriteLine Support"

#Region "WriteLine Format Support"
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
#End Region
    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal enmTarget As enmTarget)
        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        ' MH 16.05.06
        Try
            'GN 17.07.06
            'Show the errors in debugging mode, too
#if DEBUG
            Debug.WriteLine(strMessage)
#End If
#If TIME Then
            'Normal date info
            Dim strDate As String = DateTime.Now.ToString

            'GN 14.11.06
            'For speed measure issues, I need a program start time. The offset in TRACE will be set dependent
            'For measure issues, we use a fine time resolution
            Dim dblms As Double = (Now.Ticks - g_dtmStart.Ticks) * 0.000001
            strDate += " " + dblms.ToString() + " ms "
#End If

            Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
            log4net.ThreadContext.Properties("ClassName") = info.ClassName
            log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
            log4net.ThreadContext.Properties("MethodName") = info.MethodName
            log4net.ThreadContext.Properties("FileName") = info.FileName
            If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
                Try
                    'write to application log
                    LogApp.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
                End Try
            End If
            If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
                Try
                    'write to eventlog
                    LogEvent.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
                End Try
            End If
            If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
                Try
                    'write to startup log
                    LogStartup.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
                End Try
            End If

        Catch ex As Exception
            'We can't do anything else
            Trace.WriteLine("WriteLine, ERROR - writing '" & strMessage.ToString() & "' to logfile!!!!")
        End Try
    End Sub
    ''' <summary>
    ''' Writes a message to the specified events targets.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="enmTarget">trace targets</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLine(ByVal strMessage As Object, ByVal enmTarget As enmTarget)

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        ' MH 16.05.06
        Try
            'GN 17.07.06
            'Show the errors in debugging mode, too
#if DEBUG
            Debug.WriteLine(strMessage)
#End If
#If TIME Then
            'Normal date info
            Dim strDate As String = DateTime.Now.ToString

            'GN 14.11.06
            'For speed measure issues, I need a program start time. The offset in TRACE will be set dependent
            'For measure issues, we use a fine time resolution
            Dim dblms As Double = (Now.Ticks - g_dtmStart.Ticks) * 0.000001
            strDate += " " + dblms.ToString() + " ms "
#End If
            Dim info As New log4net.Core.LocationInfo(GetType(OnTrace))
            log4net.ThreadContext.Properties("ClassName") = info.ClassName
            log4net.ThreadContext.Properties("LineNumber") = info.LineNumber
            log4net.ThreadContext.Properties("MethodName") = info.MethodName
            log4net.ThreadContext.Properties("FileName") = info.FileName
            If m_blnIsAppLoggerEnabled AndAlso (enmTarget And enmTarget.AppLog) = enmTarget.AppLog Then
                Try
                    'write to application log
                    LogApp.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to AppLogger: " & GetErrorStr(ex))
                End Try
            End If
            If m_blnIsEventLoggerEnabled AndAlso (enmTarget And enmTarget.EventLog) = enmTarget.EventLog Then
                Try
                    'write to eventlog
                    LogEvent.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to EventLogger: " & GetErrorStr(ex))
                End Try
            End If
            If m_blnIsStartupLoggerEnabled AndAlso (enmTarget And enmTarget.StartupLog) = enmTarget.StartupLog Then
                Try
                    'write to startup log
                    LogStartup.Debug(strMessage)
                Catch ex As Exception
                    Trace.WriteLine("Error writing to StartupLogger: " & GetErrorStr(ex))
                End Try
            End If

        Catch ex As Exception
            'We can't do anything else
            Trace.WriteLine("WriteLine, ERROR - writing '" & strMessage.ToString() & "' to logfile!!!!")
        End Try
    End Sub
    ''' <summary>
    ''' Writes a message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLine(ByVal strMessage As String)
        WriteLine(strMessage, enmTarget.AppLog)
    End Sub
    ''' <summary>
    ''' Writes a message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level as defined by enmTraceLevel</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(strMessage, lngLevelMessage, enmTarget.AppLog)
    End Sub
    ''' <summary>
    ''' Writes a message to the AppLog.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level as defined by enmTraceLevel</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLine(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal enmTarget As enmTarget)
        WriteLineLevel(strMessage, lngLevelMessage, enmTarget)
    End Sub
    Public Shared Sub Writeline(ByVal strMessage As String, ByVal lngLevelMessage As enmTraceLevel, ByVal strFilePath As String, ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub
#End Region

#Region "WriteLineLevel Support"
#Region "WriteLineLevel Format Support"
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
#End Region
    Public Shared Sub WriteLineLevel(ByVal strMessage As String,
                     ByVal lngLevelMessage As enmTraceLevel,
                     ByVal strFilePath As String,
                     ByVal strFileName As String)
        TraceToTarget(strMessage, strFilePath, strFileName, lngLevelMessage, enmTraceType.TypeWriteLine)
    End Sub

    ''' <summary>
    ''' Writes a message to the AppLog depending on the specified trace level.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level as defined by enmTraceLevel</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLineLevel(ByVal strMessage As String,
                                     ByVal lngLevelMessage As enmTraceLevel)
        WriteLineLevel(strMessage, lngLevelMessage, enmTarget.AppLog)
    End Sub
    Public Shared Sub WriteLineLevel(ByVal strMessage As String,
                                    ByVal lngLevelMessage As enmTraceLevel,
                                    ByVal enmTarget As enmTarget)
        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        Try
            'if no level is given, the message is NOT shown!
            'Level variable must be set, if detailed info is desired!!!
            Dim bShowMessage As Boolean = False

            If lngLevelMessage > 0 Then
                If (lngLevelMessage And m_lngTraceLevel) > 0 Then
                    'the level is found
                    bShowMessage = True
                End If
            End If

            'Now show the message
            If bShowMessage Then
                WriteLine(strMessage, enmTarget)
            End If

        Catch ex As Exception
            'We cant do anything else
            Trace.WriteLine("WriteLineLevel, ERROR - writing '" & strMessage.ToString() & "' to logfile!!!!")
        End Try
    End Sub
    ''' <summary>
    ''' Writes a message to the specified events targets depending on the specified trace level.
    ''' </summary>
    ''' <param name="strMessage">message to write</param>
    ''' <param name="lngLevelMessage">trace level as defined by enmTraceLevel</param>
    ''' <param name="enmTarget">trace target</param>
    ''' <remarks>
    ''' The default message type (eg. visible in Eventlog or with log4net.Layout.SimpleLayout) is "Information".
    ''' Otherwise use TraceError or TraceWarning.
    ''' </remarks>
    Public Shared Sub WriteLineLevel(ByVal strMessage As Object,
                                     ByVal lngLevelMessage As enmTraceLevel,
                                     ByVal enmTarget As enmTarget)
        'Optional ByVal enmTarget As enmTarget = enmTarget.AppLog)

        'every public method must contain this line!
        If Not m_blnConfigured Then Configure()

        Try
            'if no level is given, the message is NOT shown!
            'Level variable must be set, if detailed info is desired!!!
            Dim bShowMessage As Boolean = False

            If lngLevelMessage > 0 Then
                If (lngLevelMessage And m_lngTraceLevel) > 0 Then
                    'the level is found
                    bShowMessage = True
                End If
            End If

            'Now show the message
            If bShowMessage Then
                WriteLine(strMessage, enmTarget)
            End If

        Catch ex As Exception
            'We cant do anything else
            Trace.WriteLine("WriteLineLevel, ERROR - writing '" & strMessage.ToString() & "' to logfile!!!!")
        End Try
    End Sub

#End Region

#Region "Utils"
    Private Shared Function GetStrFilePath(strFilePath As String, ByRef strFileName As String) As String

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
    ''' Gets all appender's names configured for the specified logger
    ''' </summary>
    ''' <param name="strLoggerName">logger name whose appender names shall be returned</param>
    ''' <returns>string() appender names</returns>
    ''' <remarks></remarks>
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
            Trace.WriteLine("GetAppenderNamesForLogger, error: " & GetErrorStr(ex))
        End Try

        Return astrValues

    End Function
    ''' <summary>
    '''     Delivers the name of the logfile: Path and filename, or name of eventlog.
    '''     In the client applications, we produce a combined path to the Documents and Settings
    '''     data directory.
    ''' </summary>
    ''' <param name="enmLogfileType" type="OnTrace.enmConfigureTarget">
    ''' </param>
    ''' <returns>
    '''     A String value...
    ''' </returns>
    ''' <history>
    '''     GN 22.10.07 Creation
    ''' </history>
    Public Shared Function GetLogFileName(ByVal enmLogfileType As enmConfigureTarget) As String

        If Not m_blnConfigured Then Configure()
        'Define the return value
        Dim strLogfileName As String = String.Empty

        Dim strAppender As String = String.Empty
        Select Case enmLogfileType

            Case enmConfigureTarget.AppLog
                strAppender = "LogFileAppender"

            Case enmConfigureTarget.EventLog
                'For this item, it's not the path, but the name of the EventLog
                strAppender = "EventlogAppender"

            Case enmConfigureTarget.StartupLog
                strAppender = "StartupLogFileAppender"
        End Select


        Try
            Dim strConfigFilename As String = AppDomain.CurrentDomain.GetData("APP_CONFIG_FILE").ToString()
            Trace.WriteLine(String.Format("GetLogFileName: {0} for Appender: {1}", strConfigFilename, strAppender))
            Dim doc As New System.Xml.XmlDocument
            doc.Load(strConfigFilename)
            If doc IsNot Nothing Then
                'Create a XPath instruction to access the appender and its file param
                Dim nodelist As Xml.XmlNodeList = doc.SelectNodes("configuration/log4net/appender[@name='" & strAppender & "']/param[@name='File']/attribute::value")
                If Not nodelist Is Nothing Then
                    'We assume only one item
                    strLogfileName = nodelist.Item(0).Value
                End If
            End If
        Catch ex As Exception
            Trace.WriteLine("GetLogFileName (Configuration for OnTrace is missing?), error: " & GetErrorStr(ex))
        End Try

        'return the directory
        Return strLogfileName
    End Function

    Public Shared Sub PatchAppenders()

        If m_blnConfigured Then
            Return
        End If

        LicenseCheck()
        Dim appenders As log4net.Appender.IAppender()
        'reset the configuration, we patch the "File" value below... and we the fresh values
        LogApp.Logger.Repository.ResetConfiguration()
        log4net.Config.XmlConfigurator.Configure()
        appenders = LogApp.Logger.Repository.GetAppenders()
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
                    'resolve environment variables... don't know if this is good here, it's a workaround .. later i think about this
                    If baseLogfilePath.Contains("%") Then
                        Try
                            Dim FirstCharIndex As Integer = baseLogfilePath.IndexOf("%")
                            Dim LastCharIndex As Integer = baseLogfilePath.LastIndexOf("%")
                            Dim finalstring As String = baseLogfilePath.Substring(FirstCharIndex + 1, (LastCharIndex - 1 - FirstCharIndex))
                            Dim envVar As String = Environment.GetEnvironmentVariable(finalstring)
                            If String.IsNullOrEmpty(envVar) Then
                                envVar = "C:\ProgramData\Online Software AG\Logs"
                            End If
                            baseLogfilePath = baseLogfilePath.Replace("%" + finalstring + "%", envVar)
                        Catch ex As Exception
                            'dont'do anything..  we create an standard log under c:\programdata
                            baseLogfilePath = "C:\ProgramData\Online Software AG\Logs"
                        End Try

                    End If

                    Dim maybeUnpatchedString As String = CType(currentAppender, log4net.Appender.FileAppender).File
                    If (maybeUnpatchedString.Contains(System.AppDomain.CurrentDomain.BaseDirectory())) Then
                        maybeUnpatchedString = maybeUnpatchedString.Replace(System.AppDomain.CurrentDomain.BaseDirectory(), "")
                    End If
                    If maybeUnpatchedString.Contains("{Client}") Or maybeUnpatchedString.Contains("{ProgramFolderName}") Or maybeUnpatchedString.Contains("{ProgramFolderName}") Or maybeUnpatchedString.Contains("{ProgramFilename}") Then
                        If (Not maybeUnpatchedString.Contains(baseLogfilePath)) Then
                            Dim patchedString As String = maybeUnpatchedString.Replace("{Client}", client).Replace("{ProgramFolderName}", folderName).Replace("{ProgramFilename}", fileName)
                            Dim resultPath As String = baseLogfilePath + patchedString
                            CType(currentAppender, log4net.Appender.FileAppender).File = resultPath
                            CType(currentAppender, log4net.Appender.FileAppender).ActivateOptions()
                        End If
                    Else
                        If Not maybeUnpatchedString.Contains(baseLogfilePath) Then
                            Dim resultPath As String = baseLogfilePath + maybeUnpatchedString
                            CType(currentAppender, log4net.Appender.FileAppender).File = resultPath
                            CType(currentAppender, log4net.Appender.FileAppender).ActivateOptions()
                        End If
                    End If

                End If

            Catch ex As Exception
                'if the initialization failed, we don't set the m_blnConfigured flag
                m_blnConfigured = False
                Return
            End Try
        Next

        'read default trace level
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

        m_blnConfigured = True
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
        Dim sb As New Text.StringBuilder
        Const OUTER_SEPERATER As String = "================================================================="
        Const INNER_SEPERATER As String = "-----------------------------------------------------------------"

        sb.Append(description)
        sb.Append(Environment.NewLine)
        sb.Append(OUTER_SEPERATER)
        sb.Append(Environment.NewLine)
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