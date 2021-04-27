Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Public Class ClsEsternalizzaStampa
    Public Const ESTERNALIZZA_CODICESOCIETA As String = "RBES"
    Public Const ESTERNALIZZA_COMMESSA As String = "FT"
    Public Const ESTERNALIZZA_COLORE As String = "NO"
    Public Const ESTERNALIZZA_CANALEDISTRIBUZIONE As String = "P"
    Public Const ESTERNALIZZA_POSTALIZZAZIONE_NORITORNO As String = "MAS"
    Public Const ESTERNALIZZA_POSTALIZZAZIONE_CONRITORNO As String = "NOS"

    Public Const MODELLO_DOCUMENTO As Integer = 0
    Public Const MODELLO_BOLLETTINO As Integer = 1
    Public Const MODELLO_CARTAINTESTATA As Integer = 2

    Public Const ORIENTAMENTO_ORIZZONTALE As Integer = 0
    Public Const ORIENTAMENTO_VERTICALE As Integer = 1

    Public Const CHRNEWLINE As String = "|"

    'Public Structure objListModelliEsternalizza
    '   Dim nTipoModello As Integer
    '   Dim nPagine As Integer
    '   Dim sOrientation As String
    '   Dim oBollettini() As objBollettino
    'End Structure

    'Public Structure objBollettino
    '   Dim TipoDocumento As String
    '   Dim sAutorizzazione As String
    '   Dim sContoCorrente As String
    '   Dim sCodIBAN As String
    '   Dim sIntestazioneConto As String
    '   Dim sAnagraficaVersante As String
    '   Dim sCFPIVAMAV As String
    '   Dim sImpBollettino As String
    '   Dim sDataScadenza As String
    '   Dim sCausale As String
    '   Dim sCodCliente As String
    '   Dim sCodBarre As String
    'End Structure
    ''' <summary>
    ''' Funzione che traduce in un file CSV con separatore<em>;</em> l'elenco dei documenti prodotti secondo la seguente struttura:
    ''' codice società; codice divisione; codice cliente; codice cliente divisione; nome; cognome; ragione sociale; indirizzo; civico; località; provincia; CAP; partita iva; codice fiscale; numero documento; data emissione documento; agente; sezionale; tipo; colore; canale distribuzione; SAP cliente; nome cliente; indirizzo cliente; autorizzazione cliente; postalizzazione; tipo postalizzazione; tipo stampa; vassoio carta; link; tipo stampa allegati; vassoio carta allegati; link allegati; flyer; barcode; priorità; orientamento; orientamento allegati; codice nazione; email 01 ... 03;filler 01 ... 15 -> diventano i dati dei bollettini
    ''' </summary>
    ''' <param name="sMyFileEsternalizza"></param>
    ''' <param name="sTipoPostalizzazione"></param>
    ''' <param name="sCodCliente"></param>
    ''' <param name="sCap"></param>
    ''' <param name="sCodiceNazione"></param>
    ''' <param name="oListModelli"></param>
    ''' <param name="sNomeDoc"></param>
    ''' <param name="sMyErr"></param>
    ''' <returns></returns>
    Public Function WriteFileEsternalizza(ByVal sMyFileEsternalizza As String, ByVal sTipoPostalizzazione As String, ByVal sCodCliente As String, ByVal sCap As String, ByVal sCodiceNazione As String, ByVal oListModelli() As objListModelliEsternalizza, ByVal sNomeDoc As String, ByRef sMyErr As String) As Boolean
        'Public Function WriteFileEsternalizza(ByVal sMyFileEsternalizza As String, ByVal sTipoPostalizzazione As String, ByVal sCodCliente As String, ByVal sCap As String, ByVal sCodiceNazione As String, ByVal oListModelli() As objListModelliEsternalizza, ByVal sNomeDoc As String, ByVal nPagesDoc As Integer, ByRef sMyErr As String) As Boolean

        Dim MyFileToWrite As IO.StreamWriter
        Dim sDatiFile As String = ""
        Dim sVassoioCarta As String = ""
        Dim sOrientamento As String = ""
        Dim sOrientamentoDoc As String = ""
        Dim sDatiBollettini As String = ""
        Dim x, y, nPages As Integer

        Try
            nPages = 0
            MyFileToWrite = IO.File.AppendText(sMyFileEsternalizza)
            For x = 0 To oListModelli.GetUpperBound(0)
                sOrientamentoDoc = ""
                Select Case oListModelli(0).nTipoModello
                    Case MODELLO_DOCUMENTO 'FRONTE/RETRO su carta bianca
                        For y = 1 To oListModelli(0).nPagine
                            If y Mod 2 = 0 Then
                                sVassoioCarta += "WB"
                            Else
                                sVassoioCarta += "WF"
                            End If
                            If IsNothing(oListModelli(0).sOrientation) Then
                                sOrientamentoDoc += ORIENTAMENTO_VERTICALE.ToString
                            End If
                            nPages += 1
                        Next
                        If sVassoioCarta.EndsWith("WF") Then
                            sVassoioCarta = sVassoioCarta.Substring(0, Len(sVassoioCarta) - 2) & "WS"
                        End If
                        If sOrientamentoDoc = "" Then
                            sOrientamento += oListModelli(0).sOrientation
                        Else
                            sOrientamento += sOrientamentoDoc
                        End If
                    Case MODELLO_BOLLETTINO 'FRONTE/RETRO su carta bianca + FRONTE/RETRO su prefincato
                        For y = 1 To oListModelli(0).nPagine
                            If y Mod 2 = 0 Then
                                sVassoioCarta += "YB"
                            Else
                                sVassoioCarta += "YF"
                            End If
                            If IsNothing(oListModelli(0).sOrientation) Then
                                sOrientamentoDoc += ORIENTAMENTO_ORIZZONTALE.ToString
                            End If
                            nPages += 1
                        Next
                        If sVassoioCarta.EndsWith("YF") Then
                            sVassoioCarta = sVassoioCarta.Substring(0, Len(sVassoioCarta) - 2) & "YS"
                        End If
                        If sOrientamentoDoc = "" Then
                            sOrientamento += oListModelli(0).sOrientation
                        Else
                            sOrientamento += sOrientamentoDoc
                        End If
                    Case MODELLO_CARTAINTESTATA
                        For y = 1 To oListModelli(0).nPagine
                            If y Mod 2 = 0 Then
                                sVassoioCarta += "PB"
                            Else
                                sVassoioCarta += "PF"
                            End If
                            If IsNothing(oListModelli(0).sOrientation) Then
                                sOrientamento += ORIENTAMENTO_VERTICALE.ToString
                            End If
                            nPages += 1
                        Next
                        If sVassoioCarta.EndsWith("PF") Then
                            sVassoioCarta = sVassoioCarta.Substring(0, Len(sVassoioCarta) - 2) & "PS"
                        End If
                        If sOrientamentoDoc = "" Then
                            sOrientamento += oListModelli(0).sOrientation
                        Else
                            sOrientamento += sOrientamentoDoc
                        End If
                    Case Else
                        Return False
                End Select
            Next
            'controllo che il documento totale abbia lo stesso numero di pagine della somma dei singoli documenti
            'If nPagesDoc <> nPages Then
            '   sMyErr = vbCrLf + "Il numero di pagine del documento singolo NON coincide con la somma delle pagine dei singoli modelli usati."
            '   Return False
            'End If
            'codice societa
            sDatiFile = ESTERNALIZZA_CODICESOCIETA + ";"
            'codice divisione
            sDatiFile += ";"
            'codice cliente 
            sDatiFile += sCodCliente + ";"
            'codice cliente divisione
            sDatiFile += ";"
            'nome
            sDatiFile += ";"
            'cognome
            sDatiFile += ";"
            'ragione sociale
            sDatiFile += ";"
            'indirizzo
            sDatiFile += ";"
            'civico
            sDatiFile += ";"
            'localita
            sDatiFile += ";"
            'provincia
            sDatiFile += ";"
            'cap
            sDatiFile += sCap.PadLeft(5, "0") + ";"
            'partita iva
            sDatiFile += ";"
            'codice fiscale
            sDatiFile += ";"
            'numero documento
            sDatiFile += ";"
            'data emissione documento
            sDatiFile += ";"
            'agente
            sDatiFile += ";"
            'sezionale
            sDatiFile += ";"
            'tipo
            sDatiFile += ";"
            'colore
            sDatiFile += ESTERNALIZZA_COLORE + ";"
            'canale distribuzione
            sDatiFile += ESTERNALIZZA_CANALEDISTRIBUZIONE + ";"
            'sap cliente
            sDatiFile += ";"
            'nome cliente
            sDatiFile += ";"
            'indirizzo cliente
            sDatiFile += ";"
            'autorizzazione cliente
            sDatiFile += ";"
            'postalizzazione
            sDatiFile += sTipoPostalizzazione + ";"
            'tipo postalizzazione
            sDatiFile += ";"
            'tipo stampa
            sDatiFile += ";"
            'vassoio carta
            sDatiFile += sVassoioCarta + ";"
            'link
            sDatiFile += sNomeDoc.Replace(".doc", ".pdf") + ";"
            'tipo stampa allegati
            sDatiFile += ";"
            'vassoio carta allegati
            sDatiFile += ";"
            'link allegati
            sDatiFile += ";"
            'flyer
            sDatiFile += ";"
            'barcode
            sDatiFile += ";"
            'priorita
            sDatiFile += ";"
            'orientamento
            sDatiFile += sOrientamento + ";"
            'orientamento allegati
            sDatiFile += ";"
            'codie nazione
            sDatiFile += sCodiceNazione + ";"
            'email 01 ... 03
            sDatiFile += ";"
            'filler 01 ... 15 -> diventano i dati dei bollettini
            sDatiFile += sDatiBollettini

            MyFileToWrite.WriteLine("L" + sDatiFile)
            MyFileToWrite.Flush()
            If Not IsNothing(oListModelli(0).oBollettini) Then
                For y = 0 To oListModelli(0).oBollettini.GetUpperBound(0)
                    If oListModelli(0).oBollettini(y).sImpBollettino <> "" Then
                        sDatiBollettini = oListModelli(0).oBollettini(y).TipoDocumento + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sAutorizzazione + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sContoCorrente + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sCodIBAN + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sIntestazioneConto + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sAnagraficaVersante + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sCFPIVAMAV + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sImpBollettino + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sDataScadenza + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sCausale + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sCodCliente + ";"
                        sDatiBollettini += oListModelli(0).oBollettini(y).sCodBarre + ";;;;;;;;;;;;;;;;;"
                        MyFileToWrite.WriteLine("B" + sDatiFile + sDatiBollettini)
                        MyFileToWrite.Flush()
                    End If
                Next
            End If

            Return True
        Catch Err As Exception
            sMyErr = "Si è verificato il seguente errore:" & vbCrLf & Err.Message
            Return False
        Finally
            MyFileToWrite.Close()
        End Try
    End Function
End Class
