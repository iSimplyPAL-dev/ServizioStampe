Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto dei percorsi del documento prodotto
    ''' </summary>
    <Serializable()>
    Public Class oggettoURL
        Public Sub New()
        End Sub
        Private _url As String
        Private _name As String
        Private _percorso As String
        Private _setupDocumento As SetupDocumento

        Public Property Url() As String
            Get
                Return _url
            End Get
            Set(ByVal Value As String)
                _url = Value
            End Set
        End Property
        Public Property Name() As String
            Get
                Return _name
            End Get
            Set(ByVal Value As String)
                _name = Value
            End Set
        End Property

        Public Property Path() As String
            Get
                Return _percorso
            End Get
            Set(ByVal Value As String)
                _percorso = Value
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
    ''' Definizione oggetto dei percorsi del gruppo di documenti prodotti
    ''' </summary>
    <Serializable()>
    Public Class GruppoURL
        Private _urlcomplessivo As oggettoURL
        Private _urlgruppi As oggettoURL()
        Private _urldocumenti As oggettoURL()


        Public Sub New()
        End Sub


        Public Property URLComplessivo() As oggettoURL
            Get
                Return _urlcomplessivo
            End Get
            Set(ByVal Value As oggettoURL)
                _urlcomplessivo = Value
            End Set
        End Property

        Public Property URLGruppi() As oggettoURL()
            Get
                Return _urlgruppi
            End Get
            Set(ByVal Value As oggettoURL())
                _urlgruppi = Value
            End Set
        End Property

        Public Property URLDocumenti() As oggettoURL()
            Get
                Return _urldocumenti
            End Get
            Set(ByVal Value As oggettoURL())
                _urldocumenti = Value
            End Set
        End Property

    End Class
End Namespace
