Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters
Imports System.Runtime.Serialization.Formatters.Soap
Imports System.Runtime.Remoting.ObjRef
Imports System.Threading
Imports System.Collections
Imports System.Configuration
Imports RIBESElaborazioneDocumentiInterface
Imports System.ServiceProcess
Imports log4net
Imports log4net.Config
Imports System.IO

''' <summary>
''' Classe di iniziazione del servizio.
''' 
''' Il servizio si occupa di produrre i documenti in formato WORD o PDF.
''' </summary>
Public Class ServizioElaborazioneDocumenti
    Inherits System.ServiceProcess.ServiceBase

    'Private components As System.ComponentModel.Container = Nothing
    Private chan As HttpChannel
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(ServizioElaborazioneDocumenti))
    'true --> quando si deve buildare il servizio
    'false --> quando si vuole lanciare in console per il debug
    Private Shared _runService As Boolean = False

    Public Sub New()
        MyBase.New()
        InitializeComponent()
    End Sub

    'UserService overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'The main entry point for the process
    <MTAThread()>
    Shared Sub Main()
        Dim ServicesToRun() As System.ServiceProcess.ServiceBase

        'More than one NT Service may run within the same process. To add
        'another service to this process, change the following line to
        'create a second service object. For example,
        '
        '  ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
        '
        If _runService = True Then
            ServicesToRun = New System.ServiceProcess.ServiceBase() {New ServizioElaborazioneDocumenti}
            System.ServiceProcess.ServiceBase.Run(ServicesToRun)
        Else
            Dim oServizio As New ServizioElaborazioneDocumenti
            oServizio.OnStart(Nothing)
            Console.WriteLine("pronti...partenza...via!")
            Console.ReadLine()
        End If

    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
        Me.ServiceName = "Service1"
    End Sub

    Protected Overrides Sub OnStart(ByVal args() As String)
        'Add code here to start your service. This method should set things
        'in motion so your service can do its work.

        Dim pathfileinfo As String = ConstSession.PathFileConfLog4Net
        Dim fileconfiglog4net As New FileInfo(pathfileinfo)
        XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)
        RegisterService()
    End Sub

    Protected Overrides Sub OnStop()
        'Add code here to perform any tear-down necessary to stop your service.
        ChannelServices.UnregisterChannel(chan)
    End Sub

    Private Shared Sub RegisterService()
        Try
            Dim props As IDictionary = New Hashtable

            '*** HTTP ***
            Dim iPortaComunicazioneHTTP As Long = ConstSession.PortaComunicazioneHTTP
            Dim serverProv As SoapServerFormatterSinkProvider = New SoapServerFormatterSinkProvider
            Dim clientProv As SoapClientFormatterSinkProvider = New SoapClientFormatterSinkProvider

            props("port") = iPortaComunicazioneHTTP '52101
            props("typeFilterLevel") = TypeFilterLevel.Full

            serverProv.TypeFilterLevel = TypeFilterLevel.Full

            Dim chan As New HttpChannel(props, clientProv, serverProv)
            ChannelServices.RegisterChannel(chan)
            '*** ***

            RemotingConfiguration.RegisterWellKnownServiceType(GetType(clsPrint), "RibesElaborazioneDocumenti.soap", WellKnownObjectMode.SingleCall)

            log.Debug("Servizio di Stampa Avviato Correttamente.")
        Catch Err As Exception
            log.Debug("Errore->", Err)
        End Try
    End Sub
End Class
