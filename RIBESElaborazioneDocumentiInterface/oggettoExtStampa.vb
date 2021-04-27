'****20110926 oggetto per file CSV esternalizzazione stampa*****'
Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto per file CSV esternalizzazione stampa
    ''' </summary>
    <Serializable()>
    Public Class oggettoExtStampa
        Public Sub New()
        End Sub

        Private _Codicecliente As String
        Private _Capco As String
        Private _Codstatonazione As String
        Private _NomeFileSingolo As String
        Private _oBollettino() As objBollettino

        Public Property Codicecliente() As String
            Get
                Return _Codicecliente
            End Get
            Set(ByVal Value As String)
                _Codicecliente = Value
            End Set
        End Property

        Public Property Capco() As String
            Get
                Return _Capco
            End Get
            Set(ByVal Value As String)
                _Capco = Value
            End Set
        End Property

        Public Property Codstatonazione() As String
            Get
                Return _Codstatonazione
            End Get
            Set(ByVal Value As String)
                _Codstatonazione = Value
            End Set
        End Property

        Public Property NomeFileSingolo() As String
            Get
                Return _NomeFileSingolo
            End Get
            Set(ByVal Value As String)
                _NomeFileSingolo = Value
            End Set
        End Property

        Public Property oBollettino() As objBollettino()
            Get
                Return _oBollettino
            End Get
            Set(ByVal Value As objBollettino())
                _oBollettino = Value
            End Set
        End Property
    End Class

End Namespace