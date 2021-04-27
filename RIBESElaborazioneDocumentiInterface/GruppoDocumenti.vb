Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto per il gruppo di documenti da produrre
    ''' </summary>
    <Serializable()>
    Public Class GruppoDocumenti

        Private _testatagruppo As oggettoTestata
        Private _oggettidastampare As oggettoDaStampareCompleto()
        '*********************************************
        '****20110926 oggetto per file CSV esternalizzazione stampa*****'
        Private _objEsternalizza As oggettoExtStampa
        Private _listModelli() As objListModelliEsternalizza
        '*********************************************
        Public Sub New()
        End Sub

        Public Property TestataGruppo() As oggettoTestata
            Get
                Return _testatagruppo
            End Get
            Set(ByVal Value As oggettoTestata)
                _testatagruppo = Value
            End Set
        End Property

        Public Property OggettiDaStampare() As oggettoDaStampareCompleto()
            Get
                Return _oggettidastampare
            End Get
            Set(ByVal Value As oggettoDaStampareCompleto())
                _oggettidastampare = Value
            End Set
        End Property

        '************************************************************
        '*** 20110927 - aggiunta gestione stampa CSV esternalizza ***
        Public Property objEsternalizza() As oggettoExtStampa
            Get
                Return _objEsternalizza
            End Get
            Set(ByVal Value As oggettoExtStampa)
                _objEsternalizza = Value
            End Set
        End Property
        '************************************************************
        '*** 20110927 - aggiunta gestione stampa CSV esternalizza ***
        Public Property listModelli() As objListModelliEsternalizza()
            Get
                Return _listModelli
            End Get
            Set(ByVal Value As objListModelliEsternalizza())
                _listModelli = Value
            End Set
        End Property
        '************************************************************
    End Class
End Namespace
