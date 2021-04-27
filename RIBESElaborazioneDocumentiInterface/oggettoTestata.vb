
Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione dati generali di configurazione da stampare
    ''' </summary>
    <Serializable()>
    Public Class oggettoTestata
        Public Sub New()
        End Sub
        Private _ente As String
        Private _atto As String
        Private _dominio As String
        Private _filename As String
        Private _setupDocumento As SetupDocumento = New SetupDocumento

        Public Property Filename() As String
            Get
                Return _filename
            End Get
            Set(ByVal Value As String)
                _filename = Value
            End Set
        End Property

        Public Property Ente() As String
            Get
                Return _ente
            End Get
            Set(ByVal Value As String)
                _ente = Value
            End Set
        End Property

        Public Property Atto() As String
            Get
                Return _atto
            End Get
            Set(ByVal Value As String)
                _atto = Value
            End Set
        End Property
        Public Property Dominio() As String
            Get
                Return _dominio
            End Get
            Set(ByVal Value As String)
                _dominio = Value
            End Set
        End Property

        Public Property oSetupDocumento() As SetupDocumento
            Get
                Return _setupDocumento
            End Get
            Set(ByVal Value As SetupDocumento)
                _setupDocumento = Value
            End Set
        End Property

    End Class

    ''' <summary>
    ''' Definizione oggetto che serve per indicare, al fine dell'unione di più documenti, che tipo di formattazione andrà applicata alla sezione che viene inserita in coda al documento aperto per l'unione
    ''' </summary>
    <Serializable()>
    Public Class SetupDocumento

        Private _Orientamento As String = "V"
        Private _MargineTOP As Integer = -1
        Private _MargineBOTTOM As Integer = -1
        Private _MargineLEFT As Integer = -1
        Private _MargineRIGHT As Integer = -1
        Private _FirstPageTray As Integer = -1
        Private _OtherPageTray As Integer = -1
        Private _PdfDoc As Boolean = False 'se true è pdf, altrimenti è doc

        Public Sub New()

        End Sub

        Public Property Orientamento() As String
            Get
                Return _Orientamento
            End Get
            Set(ByVal Value As String)
                _Orientamento = Value
            End Set
        End Property

        Public Property MargineTop() As Integer
            Get
                Return _MargineTOP
            End Get
            Set(ByVal Value As Integer)
                _MargineTOP = Value
            End Set
        End Property

        Public Property MargineBottom() As Integer
            Get
                Return _MargineBOTTOM
            End Get
            Set(ByVal Value As Integer)
                _MargineBOTTOM = Value
            End Set
        End Property

        Public Property MargineLeft() As Integer
            Get
                Return _MargineLEFT
            End Get
            Set(ByVal Value As Integer)
                _MargineLEFT = Value
            End Set
        End Property

        Public Property MargineRight() As Integer
            Get
                Return _MargineRIGHT
            End Get
            Set(ByVal Value As Integer)
                _MargineRIGHT = Value
            End Set
        End Property

        Public Property FirstPageTray() As Integer
            Get
                Return _FirstPageTray
            End Get
            Set(ByVal Value As Integer)
                _FirstPageTray = Value
            End Set
        End Property

        Public Property OtherPageTray() As Integer
            Get
                Return _OtherPageTray
            End Get
            Set(ByVal Value As Integer)
                _OtherPageTray = Value
            End Set
        End Property

        Public Property PdfDoc() As Boolean
            Get
                Return _PdfDoc
            End Get
            Set(ByVal Value As Boolean)
                _PdfDoc = Value
            End Set
        End Property

    End Class
End Namespace

