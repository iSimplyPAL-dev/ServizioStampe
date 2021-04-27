Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto da stampare con riferimenti del template
    ''' </summary>
    <Serializable()>
    Public Class oggettoDaStampareCompleto
        Private _testatadoc As oggettoTestata
        Private _testatadot As oggettoTestata
        Private _stampa As oggettiStampa()
        '*** 20101014 - aggiunta gestione stampa barcode ***
        Private _listBarcode() As ObjBarcodeToCreate

        Public Sub New()
        End Sub

        Public Property TestataDOT() As oggettoTestata
            Get
                Return _testatadot
            End Get
            Set(ByVal Value As oggettoTestata)
                _testatadot = Value
            End Set
        End Property

        Public Property TestataDOC() As oggettoTestata
            Get
                Return _testatadoc
            End Get
            Set(ByVal Value As oggettoTestata)
                _testatadoc = Value
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
        '*** 20101014 - aggiunta gestione stampa barcode ***
        Public Property oListBarcode() As ObjBarcodeToCreate()
            Get
                Return _listBarcode
            End Get
            Set(ByVal Value As ObjBarcodeToCreate())
                _listBarcode = Value
            End Set
        End Property
        '*********************************************
    End Class
    '*** 20101014 - aggiunta gestione stampa barcode ***
    <Serializable()>
    Public Class ObjBarcodeToCreate
        Dim _nType As Integer
        Dim _sBookmark As String
        Dim _sData As String

        Public Property nType() As Integer
            Get
                Return _nType
            End Get
            Set(ByVal value As Integer)
                _nType = value
            End Set
        End Property
        Public Property sBookmark() As String
            Get
                Return _sBookmark
            End Get
            Set(ByVal value As String)
                _sBookmark = value
            End Set
        End Property
        Public Property sData() As String
            Get
                Return _sData
            End Get
            Set(ByVal value As String)
                _sData = value
            End Set
        End Property
    End Class
    '*********************************************
End Namespace
