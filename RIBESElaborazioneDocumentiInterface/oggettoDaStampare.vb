Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto da stampare
    ''' </summary>
    <Serializable()>
    Public Class oggettoDaStampare
        Private _testata As oggettoTestata
        Private _stampa As oggettiStampa()


        Public Sub New()
        End Sub


        Public Property Testata() As oggettoTestata
            Get
                Return _testata
            End Get
            Set(ByVal Value As oggettoTestata)
                _testata = Value
            End Set
        End Property

        Public Property Stampa() As oggettiStampa()
            Get
                Return _stampa
            End Get
            Set(ByVal Value As oggettiStampa())
                _stampa = Value
            End Set
        End Property
    End Class
End Namespace
