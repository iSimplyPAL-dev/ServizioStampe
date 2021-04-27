Namespace Stampa.oggetti
    ''' <summary>
    ''' Definizione oggetto stampa
    ''' </summary>
    <Serializable()>
    Public Class oggettiStampa
        Public Sub New()
        End Sub

        Private _valore As String
        Private _descrizione As String
        Private _appartenenza As String
        Private _codTributo As String
        Private _ente As String
        Private _tipo As String
        Private _numFabb As String
        Private _anno As String
        '*** 20131104 - TARES ***
        Private _Sezione As String = "EL"
        Private _Rateizzazione As String = ""
        Private _IsAcconto As String = ""
        Private _IsSaldo As String = ""
        '*** ***
        Private _IsRavvedimento As String = ""

        Public Property Valore() As String
            Get
                Return _valore
            End Get
            Set(ByVal Value As String)
                _valore = Value
            End Set
        End Property

        Public Property Descrizione() As String
            Get
                Return _descrizione
            End Get
            Set(ByVal Value As String)
                _descrizione = Value
            End Set
        End Property

        Public Property Appartenenza() As String
            Get
                Return _appartenenza
            End Get
            Set(ByVal Value As String)
                _appartenenza = Value
            End Set
        End Property

        Public Property CodTributo() As String
            Get
                Return _codTributo
            End Get
            Set(ByVal Value As String)
                _codTributo = Value
            End Set
        End Property

        Public Property NumFabb() As String
            Get
                Return _numFabb
            End Get
            Set(ByVal Value As String)
                _numFabb = Value
            End Set
        End Property

        Public Property Anno() As String
            Get
                Return _anno
            End Get
            Set(ByVal Value As String)
                _anno = Value
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

        Public Property Tipo() As String
            Get
                Return _tipo
            End Get
            Set(ByVal Value As String)
                _tipo = Value
            End Set
        End Property
        '*** 20131104 - TARES **
        Public Property Sezione() As String
            Get
                Return _Sezione
            End Get
            Set(ByVal Value As String)
                _Sezione = Value
            End Set
        End Property
        Public Property Rateizzazione() As String
            Get
                Return _Rateizzazione
            End Get
            Set(ByVal Value As String)
                _Rateizzazione = Value
            End Set
        End Property
        Public Property IsAcconto() As String
            Get
                Return _IsAcconto
            End Get
            Set(ByVal Value As String)
                _IsAcconto = Value
            End Set
        End Property
        Public Property IsSaldo() As String
            Get
                Return _IsSaldo
            End Get
            Set(ByVal Value As String)
                _IsSaldo = Value
            End Set
        End Property
        '*** ***
        Public Property IsRavvedimento() As String
            Get
                Return _IsRavvedimento
            End Get
            Set(ByVal Value As String)
                _IsRavvedimento = Value
            End Set
        End Property
    End Class
    '****20110926 oggetto per file CSV esternalizzazione stampa*****'
    ''' <summary>
    ''' Definizione oggetto documenti per l'esternalizzazione stampa
    ''' </summary>
    <Serializable()>
    Public Structure objListModelliEsternalizza
        Dim nTipoModello As Integer
        Dim nPagine As Integer
        Dim sOrientation As String
        Dim oBollettini() As objBollettino
    End Structure
    '****20110926 oggetto per file CSV esternalizzazione stampa*****'
    ''' <summary>
    ''' Definizione oggetto bollettini per l'esternalizzazione stampa
    ''' </summary>
    <Serializable()>
    Public Structure objBollettino
        Dim TipoDocumento As String
        Dim sAutorizzazione As String
        Dim sContoCorrente As String
        Dim sCodIBAN As String
        Dim sIntestazioneConto As String
        Dim sAnagraficaVersante As String
        Dim sCFPIVAMAV As String
        Dim sImpBollettino As String
        Dim sDataScadenza As String
        Dim sCausale As String
        Dim sCodCliente As String
        Dim sCodBarre As String
    End Structure
End Namespace
