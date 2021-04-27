Imports System
Imports System.Configuration
Imports System.Data.SqlClient
Imports System.IO
Imports System.Threading
Imports System.Xml
Imports System.Windows.Forms
Imports System.Drawing.Printing
Imports System.Drawing
Imports RIBESElaborazioneDocumentiInterface
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti
Imports System.Collections.ArrayList
Imports log4net.Config
Imports log4net
'Imports WebSupergoo.ABCpdf7
'Imports WebSupergoo.ABCpdf7.Objects
'Imports WebSupergoo.ABCpdf7.Atoms
'Imports WebSupergoo.ABCpdf7.Operations
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports iTextSharp.text.pdf

''' <summary>
''' Classe che incapsula tutte le costanti
''' </summary>
Public Class ConstSession
    Private Shared Log As ILog = LogManager.GetLogger(GetType(ConstSession))
    ''' <summary>
    ''' 
    ''' </summary>
    Public Enum TypeAppend
        ToBegin = 1
        ToEnd = 2
    End Enum
    ''' <summary>
    ''' 
    ''' </summary>
    Public Class ManagedExtensions
        Public Const Office As String = ".doc"
        Public Const OfficeXML As String = ".dot"
        Public Const PDF As String = ".pdf"
    End Class
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PathFileConfLog4Net() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("pathfileconflog4net") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("pathfileconflog4net").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.PathFileConfLog4Net.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PortaComunicazioneHTTP() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("PortaComunicazioneHTTP") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("PortaComunicazioneHTTP").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.PortaComunicazioneHTTP.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PortaComunicazioneTCP() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("PortaComunicazioneTCP") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("PortaComunicazioneTCP").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.PortaComunicazioneTCP.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property CopyDir() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("copydir") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("copydir").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.CopyDir.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property ExtDir() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("ExtDir") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("ExtDir").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.ExtDir.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property ThreadSleep() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("ThreadSleep") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("ThreadSleep").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.ThreadSleep.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property URLWSStampaBarcode() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("URLWSStampaBarcode") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("URLWSStampaBarcode").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.URLWSStampaBarcode.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property LocationPDF() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("LocationPDF") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("LocationPDF").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.LocationPDF.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property StampantePredefinita() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("StampantePredefinita") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("StampantePredefinita").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.StampantePredefinita.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property StampantePDF() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("StampantePDF") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("StampantePDF").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.StampantePDF.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property PathCopyDoc() As String
        Get
            Try
                If (ConfigurationManager.AppSettings("PathCopyDoc") Is Nothing) Then
                    Return ""
                Else
                    Return ConfigurationManager.AppSettings("PathCopyDoc").ToString
                End If
            Catch ex As Exception
                Log.Debug("ConstSession.PathCopyDoc.errore: ", ex)
                Return ""
            End Try
        End Get
    End Property
End Class
''' <summary>
''' Classe rende disponibili le interfacce di produzione documenti 
''' </summary>
Public Class clsPrint
    Inherits MarshalByRefObject
    Implements IElaborazioneStampaDocOggetti
    Private Shared Log As ILog = LogManager.GetLogger(GetType(clsPrint))
    ''' <summary>
    ''' Per ogni gruppo stampa tutti i documenti tramite la funzione CallPrinter.
    ''' Se l'oggetto testata di gruppi di documenti in ingresso è popolato, devo unire i documenti appena generati tramite la funzione UnionDocument.
    ''' Se necessario scrivo il file per l'esternalizzazione della stampa tramite la funzione WriteFileEsternalizza.
    ''' Se necessario converto In PDF.
    ''' Se l'oggetto testata in ingresso è popolato, devo unire i gruppi di documenti generati tramite la funzione UnionDocument.
    ''' Se URLComplessivo è valorizzato cancello i singoli file che hanno composto il file complessivo e svuoto la cartella temp di appoggio.
    ''' </summary>
    ''' <param name="PathTemplate"></param>
    ''' <param name="PathVirtualTemplate"></param>
    ''' <param name="TestataGruppo"></param>
    ''' <param name="ListGruppiDoc"></param>
    ''' <param name="bIsStampaBollettino"></param>
    ''' <param name="bCreaPDF"></param>
    ''' <returns></returns>
    ''' <revisionHistory><revision date="17/03/2020">per GC il motore di stampa è su un server diverso da quello di visualizzazione dei documenti e non si riesce ad accedervi tramite IIS per questioni di autorizzazione; è stata aggiunta la copia da server a server dopo la creazione del documento</revision></revisionHistory>
    ''' <revisionHistory><revision date="19/01/2021">aggiunta chiusura forzata di tutte le istanze di word eventualmente aperte sul server</revision></revisionHistory>
    Public Function StampaDocumenti(ByVal PathTemplate As String, ByVal PathVirtualTemplate As String, ByVal TestataGruppo As oggettoTestata, ByVal ListGruppiDoc() As GruppoDocumenti, ByVal bIsStampaBollettino As Boolean, ByVal bCreaPDF As Boolean) As GruppoURL Implements RIBESElaborazioneDocumentiInterface.IElaborazioneStampaDocOggetti.StampaDocumenti
        Try
            Dim PathDocTemp As String = ""
            Dim myGruppoUrl As New GruppoURL
            Dim ListSingleDoc() As oggettoURL
            Dim DocPerGroup As New oggettoURL
            Dim ListDocPerGroup As New ArrayList
            Dim ListDocTotale As New ArrayList
            Dim fncEsternalizza As New ClsEsternalizzaStampa
            Dim sFileEsternalizzaStampa As String
            Dim ErrStampaDoc As String = ""
            Dim oFi As System.IO.FileInfo

            'KillHungProcess("WinWord.exe")

            'prendo tutti i gruppi di documenti presenti nell'array
            For Each myGruppoDocumenti As Stampa.oggetti.GruppoDocumenti In ListGruppiDoc
                Dim fncPrintDoc As New clsPrintDocument(PathTemplate, PathVirtualTemplate)
                Dim oListModelli() As objListModelliEsternalizza = Nothing

                PathDocTemp = PathTemplate
                If myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Atto.CompareTo("") <> 0 Then
                    PathDocTemp += myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Atto + "\"
                End If
                If myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Dominio.CompareTo("") <> 0 Then
                    PathDocTemp += myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Dominio + "\"
                End If
                If myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Ente.CompareTo("") <> 0 Then
                    PathDocTemp += myGruppoDocumenti.OggettiDaStampare(0).TestataDOC.Ente + "\"
                End If
                'per ogni gruppo stampa tutti i documenti 
                ListSingleDoc = fncPrintDoc.CallPrinter(myGruppoDocumenti.OggettiDaStampare, bIsStampaBollettino, oListModelli)
                '*** 20110928 per esternalizzazione ***
                If Not IsNothing(myGruppoDocumenti.objEsternalizza) Then
                    oListModelli(0).oBollettini = myGruppoDocumenti.objEsternalizza.oBollettino
                End If
                myGruppoDocumenti.listModelli = oListModelli
                '*** ***

                For Each oURL As oggettoURL In ListSingleDoc
                    If Not IsNothing(oURL) Then
                        Log.Debug("StampaWordDocument.StampaDocumentiProva.Path stampa " & oURL.Path)
                        ListDocTotale.Add(oURL)
                    Else
                        Log.Debug("StampaWordDocument.StampaDocumentiProva.errore::oURL.Path vuoto")
                        Return Nothing
                    End If
                Next
                'se l'oggetto testata di gruppi di documenti è popolato,devo unire i documenti generati
                If Not myGruppoDocumenti.TestataGruppo Is Nothing Then
                    DocPerGroup = fncPrintDoc.UnionDocument(myGruppoDocumenti.TestataGruppo, ListSingleDoc, ConstSession.TypeAppend.ToEnd)
                    ListDocPerGroup.Add(DocPerGroup)
                Else
                    Log.Debug("StampaWordDocument.StampaDocumentiProva.oGruppoDocumenti.TestataGruppo Is Nothing quindi non unisco i documenti")
                End If

                fncPrintDoc.Chiudi()
                System.Threading.Thread.Sleep(500)
                myGruppoDocumenti.listModelli = oListModelli

                '*** 20110926 - scrivo il file per l'esternalizzazione della stampa ***
                If Not IsNothing(myGruppoDocumenti.objEsternalizza) Then
                    'imposto il nome del file esternalizza
                    sFileEsternalizzaStampa = PathTemplate & myGruppoDocumenti.TestataGruppo.Dominio & "_" & myGruppoDocumenti.TestataGruppo.Ente & ".csv"
                    If fncEsternalizza.WriteFileEsternalizza(sFileEsternalizzaStampa, ClsEsternalizzaStampa.ESTERNALIZZA_POSTALIZZAZIONE_NORITORNO, myGruppoDocumenti.objEsternalizza.Codicecliente, myGruppoDocumenti.objEsternalizza.Capco, myGruppoDocumenti.objEsternalizza.Codstatonazione, myGruppoDocumenti.listModelli, myGruppoDocumenti.objEsternalizza.NomeFileSingolo & ConstSession.ManagedExtensions.PDF, ErrStampaDoc) = False Then
                        Log.Debug("Errore nella scrittura del file per l'esternalizzazione! " & myGruppoDocumenti.objEsternalizza.Codicecliente & ErrStampaDoc)
                        Return Nothing
                    End If
                End If
                '***  ***
                '*****modifica 20110927 Emanuele*****'
                If bCreaPDF = True Then
                    If DocPerGroup.Path.IndexOf(ConstSession.ManagedExtensions.Office) > 0 Or DocPerGroup.Path.IndexOf(ConstSession.ManagedExtensions.OfficeXML) > 0 Then
                        'creo i pdf singoli per ogni doc
                        Log.Debug("devo convertire::" & DocPerGroup.Path & " ::in::" & fncPrintDoc.ExtWordToPDF(DocPerGroup.Path))
                        If fncPrintDoc.EvenPage(DocPerGroup.Path, fncPrintDoc.ExtWordToPDF(DocPerGroup.Path)) = False Then
                            'non sono riuscito a convertire ci riprovo ancora una volta
                            If fncPrintDoc.EvenPage(DocPerGroup.Path, fncPrintDoc.ExtWordToPDF(DocPerGroup.Path)) = False Then
                                Return Nothing
                            End If
                        End If
                        DocPerGroup.Path = fncPrintDoc.ExtWordToPDF(DocPerGroup.Path)
                        DocPerGroup.Name = fncPrintDoc.ExtWordToPDF(DocPerGroup.Name)
                    End If
                Else
                    Log.Debug("NON devo convertire in PDF")
                End If
            Next

            '*****modifica 20110927 Emanuele*****'
            'se l'oggetto testata è popolato, devo unire i gruppi di documenti generati
            If Not TestataGruppo Is Nothing Then
                Dim fncPrintDoc As New clsPrintDocument(PathTemplate, PathVirtualTemplate)
                Log.Debug("StampaWordDocument::StampaDocumentiProva, UnionGruppiDoc")
                myGruppoUrl.URLComplessivo = fncPrintDoc.UnionDocument(TestataGruppo, CType(ListDocPerGroup.ToArray(GetType(oggettoURL)), oggettoURL()), 2)
                fncPrintDoc.Chiudi()
            End If
            System.Threading.Thread.Sleep(500)
            If Not myGruppoUrl.URLComplessivo Is Nothing Then
                If ConstSession.PathCopyDoc <> "" Then
                    Log.Debug("clsPrint.StampaDocumenti.ho altro server")
                    If Not myGruppoUrl.URLComplessivo Is Nothing Then
                        Log.Debug("clsPrint.StampaDocumenti.PathCopyDoc->" + ConstSession.PathCopyDoc + "<-")
                        Log.Debug("clsPrint.StampaDocumenti.copio da " + myGruppoUrl.URLComplessivo.Path + " a " + myGruppoUrl.URLComplessivo.Path.Replace(ConstSession.CopyDir, ConstSession.PathCopyDoc))
                        File.Copy(myGruppoUrl.URLComplessivo.Path, myGruppoUrl.URLComplessivo.Path.Replace(ConstSession.CopyDir, ConstSession.PathCopyDoc))
                    End If
                End If
                System.Threading.Thread.Sleep(500)
                'se URLComplessivo è valorizzato cancello i singoli file che hanno composto il file complessivo
                If bCreaPDF = False And Not TestataGruppo Is Nothing Then
                    For Each myUrl As oggettoURL In CType(ListDocPerGroup.ToArray(GetType(oggettoURL)), oggettoURL())
                        oFi = New System.IO.FileInfo(myUrl.Path)
                        If (oFi.Exists) Then
                            Log.Debug("elimino i singoli che hanno composto il complessivo:" & myUrl.Path)
                            oFi.Delete()
                        End If
                    Next
                End If
            End If

            'svuota la cartella temp
            If PathDocTemp.Contains("\TEMP\") Then
                Log.Debug("svuoto cartella temp")
                Dim TempDir As New System.IO.DirectoryInfo(PathDocTemp)
                For Each oFi In TempDir.GetFiles()
                    oFi.Delete()
                Next
            End If

            myGruppoUrl.URLGruppi = CType(ListDocPerGroup.ToArray(GetType(oggettoURL)), oggettoURL())
            myGruppoUrl.URLDocumenti = CType(ListDocTotale.ToArray(GetType(oggettoURL)), oggettoURL())

            Return myGruppoUrl
        Catch ex As Exception
            Log.Debug("StampaWordDocument.StampaDocumentiProva.errore::", ex)
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Routine per la chiusura forzata di un processo
    ''' </summary>
    ''' <param name="processName"></param>
    Private Sub KillHungProcess(processName As String)
        Try
            Dim psi As ProcessStartInfo = New ProcessStartInfo
            psi.Arguments = "/IM " & processName & " /F"
            psi.FileName = "taskkill"
            Dim p As Process = New Process()
            p.StartInfo = psi
            p.Start()
        Catch ex As Exception
            Log.Debug("StampaWordDocument.KillHungProcess.errore::", ex)
        End Try
    End Sub
End Class
''' <summary>
''' Classe per la produzione dei documenti in formato DOC o PDF
''' </summary>
Public Class clsPrintDocument
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(clsPrintDocument))

    Dim pathfileinfo As String = ConstSession.PathFileConfLog4Net
    Dim fileconfiglog4net As System.IO.FileInfo = New FileInfo(pathfileinfo)

    Private WordApp As New Word.Application
    Private ArrayListPath As ArrayList = New ArrayList
    Private myPathTemplate As String = ""
    Private myPathVirtualTemplate As String = ""

    Private objFalse As Object = False

    Public Sub New()
        XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)
    End Sub
    Public Sub New(ByVal PathTemplate As String, ByVal PathVirtualTemplate As String)
        XmlConfigurator.ConfigureAndWatch(fileconfiglog4net)
        myPathTemplate = PathTemplate
        myPathVirtualTemplate = PathVirtualTemplate
    End Sub
    ''' <summary>
    ''' Traduce l'oggetto in ingresso in documento Word o PDF in base al tipo tramite le funzioni PrintWord e PrintPDF
    ''' </summary>
    ''' <param name="ListDaStampare"></param>
    ''' <param name="bIsStampaBollettino"></param>
    ''' <param name="oArrModelli"></param>
    ''' <returns></returns>
    Public Function CallPrinter(ByVal ListDaStampare As oggettoDaStampareCompleto(), ByVal bIsStampaBollettino As Boolean, ByRef oArrModelli() As objListModelliEsternalizza) As oggettoURL()
        Try
            Dim ListDoc As New ArrayList
            Dim countModelli As Integer = 0

            If ListDaStampare.Length = 0 Then
                log.Debug("CallPrinter.Array oggetti da stampare vuoto!!!")
            End If

            log.Debug(ListDaStampare.Length)
            countModelli = 0

            For Each oOggCompleto As oggettoDaStampareCompleto In ListDaStampare
                log.Debug("CallPrinter.Entro StampaDoc")
                ReDim Preserve oArrModelli(countModelli)
                Dim oModello As New objListModelliEsternalizza
                Dim myDoc As New oggettoURL
                If (oOggCompleto.TestataDOT.oSetupDocumento.PdfDoc) Then
                    myDoc = PrintPDF(oOggCompleto, bIsStampaBollettino, oModello)
                Else
                    myDoc = PrintWord(oOggCompleto, bIsStampaBollettino, oModello)
                End If

                oArrModelli(countModelli) = oModello
                log.Debug("CallPrinter.Uscito da StampaDoc")

                If Not myDoc.Url Is Nothing Then
                    ListDoc.Add(myDoc)
                    countModelli += 1
                End If
            Next

            log.Debug("CallPrinter.Documenti stampati con successo")
            Return CType(ListDoc.ToArray(GetType(oggettoURL)), oggettoURL())
        Catch ex As Exception
            log.Debug("CallPrinter.Errore::", ex)
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Funzione che riunisce in un solo documento tutti gli oggetti in ingresso.
    ''' Devo testare se ho almeno un elemento di tipo PDF, se si allora converto tutto in PDF ed unisco altrimenti unisco direttamente da Word.
    ''' </summary>
    ''' <param name="TestataGruppo"></param>
    ''' <param name="oDocDaUnire"></param>
    ''' <param name="nAccoda"></param>
    ''' <returns></returns>
    Public Function UnionDocument(ByVal TestataGruppo As oggettoTestata, ByVal oDocDaUnire As oggettoURL(), ByVal nAccoda As Integer) As oggettoURL
        Try
            Dim myURLRet As New oggettoURL
            Dim AttoTestata As String = ""
            Dim DominioTestata As String = ""
            Dim EnteTestata As String = ""
            Dim FileNameTestata As String = ""
            Dim PathFileTestataDOC As String = myPathTemplate
            Dim PathFileTestataWEB As String = myPathVirtualTemplate

            If Not TestataGruppo Is Nothing Then
                AttoTestata = TestataGruppo.Atto
                DominioTestata = TestataGruppo.Dominio
                EnteTestata = TestataGruppo.Ente
                FileNameTestata = TestataGruppo.Filename
            End If

            Try
                If AttoTestata <> "" Then
                    PathFileTestataDOC += AttoTestata + "\"
                    PathFileTestataWEB += AttoTestata + "/"
                End If
                If DominioTestata <> "" Then
                    PathFileTestataDOC += DominioTestata + "\"
                    PathFileTestataWEB += DominioTestata + "/"
                End If
                If EnteTestata <> "" Then
                    PathFileTestataDOC += EnteTestata + "\"
                    PathFileTestataWEB += EnteTestata + "/"
                End If
            Catch ex As Exception
                log.Debug("UnionDocument.errore su val path")
            End Try
            Try
                CreateDir(PathFileTestataDOC)
            Catch ex As Exception
                log.Debug("UnionDocument.errore su createdir")
            End Try
            '*** 20140509 - TASI ***
            Try
                Dim myTicks As String = DateTime.Now.ToString("ddMMyyyyHHmmss") + DateTime.Now.Ticks.ToString()
                FileNameTestata = FileNameTestata.Replace("MYTICKS", myTicks)
                '*** ***
                log.Debug("UnionDocument.FileNameTestata::" & FileNameTestata)
            Catch ex As Exception
                log.Debug("UnionDocument.errore su ticks di filenametestata")
                Return Nothing
            End Try
            '*** 20130114 - devo testare se ho almeno un elemento di tipo pdf allora converto tutto in pdf ed unisco altrimenti unisco direttamente da word ***
            Dim bHasPDF As Boolean = False
            If Not oDocDaUnire Is Nothing Then
                For Each item As oggettoURL In oDocDaUnire
                    If Not item Is Nothing Then
                        If (Path.GetExtension(item.Path).IndexOf("pdf") > 0) Then
                            bHasPDF = True
                        End If
                    Else
                        log.Debug("UnionDocument.errore:: item null in oDocDaUnire")
                        Return Nothing
                    End If
                Next
                If bHasPDF Then
                    FileNameTestata += ConstSession.ManagedExtensions.PDF
                Else
                    FileNameTestata += ConstSession.ManagedExtensions.Office
                End If
                PathFileTestataDOC += FileNameTestata
                PathFileTestataWEB += FileNameTestata
                If bHasPDF = True Then
                    log.Debug("UnionDocument.devo unire in pdf")
                    If UnionPDF(PathFileTestataDOC, oDocDaUnire) = False Then
                        Return Nothing
                    End If
                Else
                    log.Debug("UnionDocument.devo unire in word ")
                    If UnionDoc(PathFileTestataDOC, oDocDaUnire, nAccoda, bHasPDF) = False Then
                        Return Nothing
                    End If
                End If

                myURLRet.Name = FileNameTestata
                myURLRet.Path = PathFileTestataDOC
                myURLRet.Url = PathFileTestataWEB
                myURLRet.oSetupDocumento = oDocDaUnire(0).oSetupDocumento
            Else
                log.Debug("UnionDocument.errore oDocDaUnire null")
                myURLRet = Nothing
            End If
            Return myURLRet
        Catch ex As Exception
            log.Debug("UnionDocument.errore", ex)
            Return Nothing
        End Try
    End Function
    Public Sub Chiudi()
        Try
            WordApp.Quit(objFalse, objFalse, objFalse)
            Runtime.InteropServices.Marshal.FinalReleaseComObject(WordApp)
            log.Debug("Chiudi.quit da word")
        Catch ex As Exception
            log.Debug("Chiudi.errore::", ex)
        End Try
    End Sub
    ''' <summary>
    ''' Copio il template nel percorso di destinazione di appoggio.
    ''' Ciclo tutti i bookmark dell'oggetto da stampare e popolo il documento tramite funzione FillBookmark.
    ''' Se necessario stampa barcode tramite la funzione PrintBarcode.
    ''' Salvo e chiudo il documento.
    ''' </summary>
    ''' <param name="oggetto"></param>
    ''' <param name="bIsStampaBollettino"></param>
    ''' <param name="ArrModelli"></param>
    ''' <returns></returns>
    Private Function PrintWord(ByVal oggetto As oggettoDaStampareCompleto, ByVal bIsStampaBollettino As Boolean, ByRef ArrModelli As objListModelliEsternalizza) As oggettoURL
        Try
            Dim myURLRet As New oggettoURL
            Dim NameDOC As String

            'COMPONGO IL NOME DEL FILE TEMPLATE DA PRENDERE PER GENERARE IL DOCUMENTO
            Dim PathNameTemplate As String = ConstSession.CopyDir
            If oggetto.TestataDOT.Atto.CompareTo("") <> 0 Then
                PathNameTemplate += oggetto.TestataDOT.Atto + "\"
            End If
            If oggetto.TestataDOT.Dominio.CompareTo("") <> 0 Then
                PathNameTemplate += oggetto.TestataDOT.Dominio + "\"
            End If
            If oggetto.TestataDOT.Ente.CompareTo("") <> 0 Then
                PathNameTemplate += oggetto.TestataDOT.Ente + "\"
            End If
            If oggetto.TestataDOT.Filename.CompareTo("") <> 0 Then
                PathNameTemplate += oggetto.TestataDOT.Filename
            End If

            log.Debug("PrintWord.PathFileTemplate " & PathNameTemplate)

            Dim PathFileDOC As String = ConstSession.CopyDir
            Dim PathFileDOT As String = ConstSession.CopyDir
            Dim PathFileWEB As String = ConstSession.ExtDir
            If oggetto.TestataDOC.Atto.CompareTo("") <> 0 Then
                PathFileDOC += oggetto.TestataDOC.Atto + "\"
                PathFileWEB += oggetto.TestataDOC.Atto + "/"
            End If
            If oggetto.TestataDOC.Dominio.CompareTo("") <> 0 Then
                PathFileDOC += oggetto.TestataDOC.Dominio + "\"
                PathFileWEB += oggetto.TestataDOC.Dominio + "/"
            End If
            If oggetto.TestataDOC.Ente.CompareTo("") <> 0 Then
                PathFileDOC += oggetto.TestataDOC.Ente + "\"
                PathFileWEB += oggetto.TestataDOC.Ente + "/"
            End If
            CreateDir(PathFileDOC)
            log.Debug("PrintWord.Metto il ticks")
            PathFileDOT = PathFileDOC
            NameDOC = oggetto.TestataDOC.Filename + DateTime.Now.ToString("ddMMyyyyHHmmss") + DateTime.Now.Ticks.ToString() + ConstSession.ManagedExtensions.Office

            If NameDOC.CompareTo("") <> 0 Then
                PathFileDOC += NameDOC
                PathFileWEB += NameDOC
            End If

            'COPIO IL TEMPLATE NEL PERCORSO DI DESTINAZIONE
            File.Copy(PathNameTemplate, PathFileDOC)

            log.Debug("PrintWord.Template copiato - PathFileDOC=" & PathFileDOC & " - PathNameTemplate=" & PathNameTemplate & " - NameDOC=" & NameDOC & " - PathFileWEB=" & PathFileWEB)

            If NameDOC.IndexOf("Bollettino") > 0 And bIsStampaBollettino = False Then
                myURLRet.Name = NameDOC
                myURLRet.Path = PathFileDOC
                myURLRet.Url = Nothing
                myURLRet.oSetupDocumento = oggetto.TestataDOT.oSetupDocumento
            Else
                Dim ThreadSleep As String = ConstSession.ThreadSleep
                If ThreadSleep = "" Then ThreadSleep = "500"

                System.Threading.Thread.Sleep(Int(ThreadSleep))

                System.Windows.Forms.Application.DoEvents()
                log.Debug("PrintWord.Apertura Documento " + PathFileDOC)
                Dim aDoc As New Word.Document
                log.Debug("PrintWord.Word.Document open")
                'aDoc = WordApp.Documents.Open(PathFileDOC)
                Dim myDocuments As Word.Documents
                myDocuments = WordApp.Documents
                aDoc = myDocuments.Open(PathFileDOC)
                log.Debug("PrintWord.Documento " + PathFileDOC + " aperto")
                If aDoc Is Nothing Then
                    log.Debug("MUCCA")
                End If
                aDoc.Activate()
                log.Debug("PrintWord.activate")
                'CICLO TUTTI I BOOKMARK DELL'OGGETTODASTAMPARE
                For Each myBookmark As oggettiStampa In oggetto.Stampa
                    If myBookmark.Appartenenza = "totale" Then
                        If (myBookmark.CodTributo <> "detrazione") Then
                            FillBookmark(myBookmark.Descrizione, myBookmark.Valore, myBookmark.CodTributo, aDoc)
                        End If
                    Else
                        FillBookmark(myBookmark.Descrizione, myBookmark.Valore, "", aDoc)
                    End If
                Next
                log.Debug("PrintWord.popolato tutti i segnalibri")
                ArrModelli.nTipoModello = ClsEsternalizzaStampa.MODELLO_DOCUMENTO
                ArrModelli.nPagine = CInt(WordApp.ActiveDocument.ComputeStatistics(Word.WdStatistic.wdStatisticPages, IncludeFootnotesAndEndnotes:=False))
                WordApp.ActiveDocument.Save()
                log.Debug("PrintWord.Documento " + PathFileDOC + " salvato")
                WordApp.ActiveDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges, False, False)
                log.Debug("PrintWord.Documento " + NameDOC + " chiuso")
                Chiudi()

                '*** 20101014 - aggiunta gestione stampa barcode ***
                If Not IsNothing(oggetto.oListBarcode) Then
                    Dim FncPrintBarcode As New WSPrintBarcode.ServiceStampaBarcode
                    Dim x As Integer
                    'setto l'url del servizio
                    FncPrintBarcode.Url = ConstSession.URLWSStampaBarcode
                    log.Debug("PrintWord.Richiamo il WEB SERVICE Barcode all'url:: " & FncPrintBarcode.Url)
                    For x = 0 To UBound(oggetto.oListBarcode)
                        If FncPrintBarcode.PrintBarcode(oggetto.oListBarcode(x).nType, oggetto.oListBarcode(x).sData, PathFileDOC.Replace(NameDOC, ""), NameDOC, oggetto.oListBarcode(x).sBookmark) = False Then
                            log.Debug("PrintWord.Stampa Barcode.Errore!! Popolamento Barcode tipo::" & oggetto.oListBarcode(x).nType & "::da codificare::" & oggetto.oListBarcode(x).sData & "::documento::" & PathFileDOC & NameDOC)
                            Return Nothing
                        End If
                    Next
                End If
                '*********************************************
                myURLRet.Name = NameDOC
                myURLRet.Path = PathFileDOC
                myURLRet.Url = PathFileWEB
                myURLRet.oSetupDocumento = oggetto.TestataDOT.oSetupDocumento
            End If
            Return myURLRet
        Catch ex As Exception
            log.Debug("PrintWord.errore::", ex)
            Return Nothing
        End Try
    End Function
    ''' <summary>
    ''' Funzione per il popolamento di un segnalibro: viene selezionato il segnalibro in questione e viene riempito con il testo in ingresso
    ''' </summary>
    ''' <param name="BookmarkName"></param>
    ''' <param name="BookmarkValue"></param>
    ''' <param name="CodTributo"></param>
    ''' <param name="myWordDoc"></param>
    Private Sub FillBookmark(ByVal BookmarkName As String, ByVal BookmarkValue As String, ByVal CodTributo As String, ByVal myWordDoc As Word.Document)
        Try
            If CodTributo <> "" Then
                BookmarkName = "I_dovuto_" + CodTributoToSigla(CodTributo)
            End If
            myWordDoc.Bookmarks.ShowHidden = False
            WordApp.ActiveDocument.Bookmarks.Item(BookmarkName).Select()
            WordApp.Selection.Text = BookmarkValue
        Catch
            log.Debug("FillBookmark.Bookmark non trovato:: " & BookmarkName)
        End Try
    End Sub
    Private Function CodTributoToSigla(ByVal valore As String) As String
        Select Case valore
            Case "3912"
                Return "AB_PR"
            Case "3913"
                Return "FABR"
            Case "3914"
                Return "TE_AG"
            Case "3915"
                Return "TE_AG_Sta"
            Case "3916"
                Return "AR_FA"
            Case "3917"
                Return "AR_FA_Sta"
            Case "3918"
                Return "AL_FA"
            Case "3919"
                Return "AL_FA_Sta"
            Case "detrazioneDoc"
                Return "DETRAZ"
            Case "dovutaTotale"
                Return "totale"
            Case Else
                Return ""
        End Select
    End Function
    ''' <summary>
    ''' Funzione che unisce i documenti in base al formato, se PDF tramite funzione UnionPDF altrimenti tramite funzione UnionWord
    ''' </summary>
    ''' <param name="PathFileTestataDOC"></param>
    ''' <param name="oDocDaUnire"></param>
    ''' <param name="TypeAppendDoc"></param>
    ''' <param name="bCreaPDF"></param>
    ''' <returns></returns>
    Private Function UnionDoc(ByVal PathFileTestataDOC As String, ByVal oDocDaUnire As oggettoURL(), ByVal TypeAppendDoc As Integer, bCreaPDF As Boolean) As Boolean
        'devo convertire in PDF perché l'accodamento in WORD sballa la formattazione
        Try
            For Each myDoc As oggettoURL In oDocDaUnire
                Dim OutputFile As String = myDoc.Path
                If bCreaPDF Then
                    OutputFile = ExtWordToPDF(myDoc.Path)
                    myDoc.Path = ExtWordToPDF(myDoc.Path)
                    myDoc.Name = ExtWordToPDF(myDoc.Name)
                    myDoc.Url = ExtWordToPDF(myDoc.Url)
                End If
                EvenPage(myDoc.Path, OutputFile)
            Next
            If bCreaPDF Then
                UnionPDF(PathFileTestataDOC, oDocDaUnire)
            Else
                UnionWord(PathFileTestataDOC, oDocDaUnire, TypeAppendDoc)
            End If
            Return True
        Catch ex As Exception
            log.Debug("UnionDoc.errore::", ex)
            Return False
        End Try
    End Function
    ''' <summary>
    ''' Accodo i documenti in ingresso in un unico file in base al tipo di accodamento definito.
    ''' </summary>
    ''' <param name="PathFileTestataDOC"></param>
    ''' <param name="oDocDaUnire"></param>
    ''' <param name="TypeAppendDoc"></param>
    ''' <returns></returns>
    Private Function UnionWord(ByVal PathFileTestataDOC As String, ByVal oDocDaUnire As oggettoURL(), ByVal TypeAppendDoc As Integer) As Boolean
        Try
            Dim oFi As New System.IO.FileInfo(PathFileTestataDOC)
            If Not (oFi.Exists) Then
                TypeAppendDoc = ConstSession.TypeAppend.ToBegin
            End If

            For Each myDoc As oggettoURL In oDocDaUnire
                Dim MSdoc As New Word.ApplicationClass
                Try
                    MSdoc.Visible = True
                    MSdoc.Application.Visible = True
                    MSdoc.WindowState = WdWindowState.wdWindowStateMaximize
                    If TypeAppendDoc = ConstSession.TypeAppend.ToEnd Then
                        log.Debug("UnionWord.open toend")
                        MSdoc.Documents.Open(PathFileTestataDOC)
                        MSdoc.Selection.EndOf(Unit:=Word.WdUnits.wdStory)
                        MSdoc.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                        If myDoc.oSetupDocumento.Orientamento = "O" Then
                            MSdoc.Selection.Sections.Last.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                        Else
                            MSdoc.Selection.Sections.Last.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                        End If
                        If myDoc.oSetupDocumento.MargineBottom <> -1 Then
                            MSdoc.Selection.PageSetup.BottomMargin = Single.Parse((myDoc.oSetupDocumento.MargineBottom / 100).ToString())
                        End If
                        If myDoc.oSetupDocumento.MargineTop <> -1 Then
                            MSdoc.Selection.PageSetup.TopMargin = Single.Parse((myDoc.oSetupDocumento.MargineTop / 100).ToString())
                        End If
                        If myDoc.oSetupDocumento.MargineLeft <> -1 Then
                            MSdoc.Selection.PageSetup.LeftMargin = Single.Parse((myDoc.oSetupDocumento.MargineLeft / 100).ToString())
                        End If
                        If myDoc.oSetupDocumento.MargineRight <> -1 Then
                            MSdoc.Selection.PageSetup.RightMargin = Single.Parse((myDoc.oSetupDocumento.MargineRight / 100).ToString())
                        End If
                        If myDoc.oSetupDocumento.FirstPageTray <> -1 Then
                            MSdoc.Selection.PageSetup.FirstPageTray = myDoc.oSetupDocumento.FirstPageTray
                        End If
                        If myDoc.oSetupDocumento.OtherPageTray <> -1 Then
                            MSdoc.Selection.PageSetup.OtherPagesTray = myDoc.oSetupDocumento.OtherPageTray
                        End If
                        MSdoc.Selection.InsertFile(myDoc.Path)
                    Else
                        log.Debug("UnionWord.open")
                        MSdoc.Documents.Open(myDoc.Path)
                        TypeAppendDoc = ConstSession.TypeAppend.ToEnd
                    End If
                    If PathFileTestataDOC.EndsWith(ConstSession.ManagedExtensions.PDF) Then
                        MSdoc.ActiveDocument.SaveAs(PathFileTestataDOC, WdSaveFormat.wdFormatPDF)
                    ElseIf PathFileTestataDOC.EndsWith(ConstSession.ManagedExtensions.Office) Then
                        MSdoc.ActiveDocument.SaveAs(PathFileTestataDOC, WdSaveFormat.wdFormatDocument97)
                    Else
                        MSdoc.ActiveDocument.SaveAs(PathFileTestataDOC, WdSaveFormat.wdFormatDocument)
                    End If
                Catch ex As Exception
                    log.Debug("UnionWord.errore::", ex)
                    Return False
                Finally
                    If Not MSdoc Is Nothing Then
                        MSdoc.Documents.Close(objFalse, objFalse, objFalse)
                    End If
                    MSdoc.Quit(objFalse, objFalse, objFalse)
                    Runtime.InteropServices.Marshal.FinalReleaseComObject(MSdoc)
                    log.Debug("UnionWord.quit da word")
                End Try
            Next
            Return True
        Catch ex As Exception
            log.Debug("UnionWord.errore::", ex)
            Return False
        End Try
    End Function
    Public Shared Sub CreateDir(ByVal DirPath As String)
        Dim numfile1 As Integer
        Dim NewPathName As String = ""
        Dim CurrentDir As String
        Dim NextDir As String
        Dim SearchDir As String
        Dim NStart As Short
        Dim NStop As Short
        Dim DirFound As Boolean
        Dim Server As String = ""

        Try
            log.Debug("CreateDir:" & DirPath)
            If Left(DirPath, 2) = "\\" Then
                NStart = 3
                NStop = CShort(InStr(NStart, DirPath, "\"))
                Server = "\\" & Mid(DirPath, NStart, NStop - NStart) & "\"
                NStart = NStop
                DirPath = Mid(DirPath, NStart + 1, Len(DirPath))
                NStart = 1
                NStop = CShort(InStr(NStart, DirPath, "\"))
            Else
                NStart = 1
                NStop = CShort(InStr(NStart, DirPath, "\"))
            End If
            numfile1 = FreeFile()

            Do While NStop < Len(DirPath)
                If Server <> "" Then
                    CurrentDir = Mid(DirPath, NStart, NStop - NStart) & "\"
                Else
                    CurrentDir = Mid(DirPath, NStart, NStop - NStart) & "\"
                End If
                NewPathName = NewPathName & CurrentDir
                NStop = CShort(NStop + 1)
                NStart = NStop
                NextDir = Mid(DirPath, NStop, InStr(NStop, DirPath, "\") - NStop)
                If Server <> "" Then
                    SearchDir = Dir(Server & NewPathName, FileAttribute.Directory)
                Else
                    SearchDir = Dir(NewPathName, FileAttribute.Directory)
                End If
                Do While SearchDir <> ""
                    If UCase(SearchDir) = UCase(NextDir) Then
                        DirFound = True
                        Exit Do
                    End If
                    SearchDir = Dir()
                Loop
                If DirFound = False Then
                    If Server <> "" Then
                        MkDir(Server & NewPathName & NextDir & "\")
                    Else
                        MkDir(NewPathName & NextDir & "\")
                    End If
                Else
                    DirFound = False
                End If
                NStop = CShort(InStr(NStart, DirPath, "\"))
            Loop
        Catch ex As Exception
            log.Debug("Errore in CreaterDir", ex)
        End Try
    End Sub
    Public Function ExtWordToPDF(OrgFile As String) As String
        Dim RetFile As String = OrgFile
        Try
            RetFile = OrgFile.Replace(ConstSession.ManagedExtensions.Office, ConstSession.ManagedExtensions.PDF).Replace(ConstSession.ManagedExtensions.OfficeXML, ConstSession.ManagedExtensions.PDF)
        Catch ex As Exception
            log.Debug("ExtWordToPDF.errore::", ex)
        End Try
        Return RetFile
    End Function
    Public Function EvenPage(ByVal inputFile As String, ByVal outPath As String) As Boolean
        Dim MSdoc As New Word.ApplicationClass

        Try
            'se ho pagine dispari aggiungo pagina bianca
            log.Debug("EvenPage.controllo pagine pari")
            log.Debug("EvenPage.inputFile->" + inputFile + "      EvenPage.outPath->" + outPath)
            MSdoc.Documents.Open(inputFile)
            If (CInt(MSdoc.ActiveDocument.ComputeStatistics(Word.WdStatistic.wdStatisticPages, IncludeFootnotesAndEndnotes:=False)) Mod 2) <> 0 Then
                MSdoc.Selection.EndOf(Unit:=Word.WdUnits.wdStory) 'mi porto alla fine della doc
                MSdoc.Selection.InsertBreak(Type:=Word.WdBreakType.wdSectionBreakNextPage) 'inserisco un'interruzione di sezione pagina successiva
                MSdoc.ActiveDocument.Save()
            End If
            MSdoc.Documents.Close(objFalse, objFalse, objFalse)
            MSdoc.Visible = False
            MSdoc.Documents.Open(inputFile)
            MSdoc.Application.Visible = False
            MSdoc.WindowState = WdWindowState.wdWindowStateMinimize
            If outPath.EndsWith(ConstSession.ManagedExtensions.PDF) Then
                MSdoc.ActiveDocument.SaveAs(outPath, WdSaveFormat.wdFormatPDF)
            ElseIf outPath.EndsWith(ConstSession.ManagedExtensions.Office) Then
                MSdoc.ActiveDocument.SaveAs(outPath, WdSaveFormat.wdFormatDocument97)
            Else
                MSdoc.ActiveDocument.SaveAs(outPath, WdSaveFormat.wdFormatDocument)
            End If
            Return True
        Catch ex As Exception
            log.Debug("EvenPage::si è verificato il seguente errore::", ex)
            Return False
        Finally
            If Not MSdoc Is Nothing Then
                MSdoc.Documents.Close(objFalse, objFalse, objFalse)
            End If
            MSdoc.Quit(objFalse, objFalse, objFalse)
            Runtime.InteropServices.Marshal.FinalReleaseComObject(MSdoc)
            log.Debug("EvenPage.quit da word")
        End Try
    End Function
#Region "iTextSharp"
    ''' <summary>
    ''' Copio il template nel percorso di destinazione di appoggio.
    ''' Ciclo tutti i bookmark dell'oggetto da stampare e popolo il documento. 
    ''' Salvo e chiudo il documento.
    ''' </summary>
    ''' <param name="oggetto"></param>
    ''' <param name="bIsStampaBollettino"></param>
    ''' <param name="ArrModelli"></param>
    ''' <returns></returns>
    Private Function PrintPDF(ByVal oggetto As oggettoDaStampareCompleto, ByVal bIsStampaBollettino As Boolean, ByRef ArrModelli As objListModelliEsternalizza) As oggettoURL
        Try
            Dim oURLRet As New oggettoURL

            Dim AttoDOT As String = oggetto.TestataDOT.Atto
            Dim DominioDOT As String = oggetto.TestataDOT.Dominio
            Dim EnteDOT As String = oggetto.TestataDOT.Ente
            Dim FileNameDOT As String = oggetto.TestataDOT.Filename

            'per la gestione del file DOC
            Dim AttoDOC As String = oggetto.TestataDOC.Atto
            Dim DominioDOC As String = oggetto.TestataDOC.Dominio
            Dim EnteDOC As String = oggetto.TestataDOC.Ente
            Dim FileNameDOC As String = oggetto.TestataDOC.Filename

            Dim DataOra As String = DateTime.Now.ToString("ddMMyyyyHHmmss")

            'COMPONGO IL NOME DEL FILE TEMPLATE DA PRENDERE PER GENERARE IL DOCUMENTO
            Dim PathNameTemplate As String = ConstSession.CopyDir
            If AttoDOT.CompareTo("") <> 0 Then
                PathNameTemplate += AttoDOT + "\"
            End If
            If DominioDOT.CompareTo("") <> 0 Then
                PathNameTemplate += DominioDOT + "\"
            End If
            If EnteDOT.CompareTo("") <> 0 Then
                PathNameTemplate += EnteDOT + "\"
            End If
            If FileNameDOT.CompareTo("") <> 0 Then
                PathNameTemplate += FileNameDOT
            End If

            log.Debug("PathFileTemplate " & PathNameTemplate)

            'COMPONGO IL NOME DEL FILE DOC DA GENERARE E IL PERCORSO WEB
            Dim PathFileDOC As String = ConstSession.CopyDir
            Dim PathFileWEB As String = ConstSession.ExtDir
            Dim PathNameDOC As String = ""
            Dim PathNameWEB As String = ""
            If AttoDOC.CompareTo("") <> 0 Then
                PathFileDOC += AttoDOC + "\"
                PathFileWEB += AttoDOC + "/"
            End If
            If DominioDOC.CompareTo("") <> 0 Then
                PathFileDOC += DominioDOC + "\"
                PathFileWEB += DominioDOC + "/"
            End If
            If EnteDOC.CompareTo("") <> 0 Then
                PathFileDOC += EnteDOC + "\"
                PathFileWEB += EnteDOC + "/"
            End If
            CreateDir(PathFileDOC)
            log.Debug("Metto il ticks")
            'PathFileDOT = PathFileDOC
            FileNameDOC = FileNameDOC + DataOra + DateTime.Now.Ticks.ToString() + ConstSession.ManagedExtensions.PDF

            If FileNameDOC.CompareTo("") <> 0 Then
                PathNameDOC = PathFileDOC + FileNameDOC
                PathNameWEB = PathFileWEB + FileNameDOC
            End If

            log.Debug("PathFileDOC " & PathNameDOC)
            log.Debug("PathFileWEB " & PathNameWEB)

            log.Debug("Template copiato")
            log.Debug("PathFileDOC=" & PathFileDOC)
            log.Debug("PathFileTemplate=" & PathNameTemplate)
            log.Debug("FileNameDOC=" & FileNameDOC)
            log.Debug("PathFileWEB=" & PathFileWEB)

            'COPIO IL TEMPLATE NEL PERCORSO DI DESTINAZIONE
            Dim PDFStamper As New PdfStamper(New PdfReader(PathNameTemplate), New FileStream(PathNameDOC, FileMode.Create))

            Dim pdfFormFields As AcroFields
            Dim Array() As String
            Dim Scope As String
            Scope = ""
            'Ciclo sui segnalibri da stampare e li popolo
            For Each objBookMark As oggettiStampa In oggetto.Stampa
                pdfFormFields = PDFStamper.AcroFields
                log.Debug("pdfFormFields")
                If (objBookMark.Appartenenza <> "") Then
                    log.Debug("PrintPDF.ho appartenenza->" + objBookMark.Appartenenza + " ho valore->" + objBookMark.Valore)
                    Array = objBookMark.Valore.ToString().Split(",")
                    If objBookMark.Descrizione.IndexOf("T_DEBITI") >= 0 Then
                        Scope = objBookMark.Appartenenza
                    End If
                    If (objBookMark.CodTributo = "dovutaTotale") Then
                        If Scope = objBookMark.Appartenenza Then
                            log.Debug("PrintPDF.ho scope=appartenenza->" + Scope + " splitto->" + objBookMark.Valore + " per ,")
                            Array = objBookMark.Valore.ToString().Split(", ")
                            pdfFormFields.SetField(objBookMark.Descrizione, Array(0))
                            pdfFormFields.SetField(objBookMark.Descrizione + "b", Array(0))
                            pdfFormFields.SetField("T_TOTSALDODEC_SD", Array(1))
                            pdfFormFields.SetField("T_TOTSALDODEC_SDb", Array(1))
                        End If
                    ElseIf (objBookMark.CodTributo = "detrazione" Or objBookMark.Descrizione.IndexOf("IDOPERAZ") > 0) Then
                        pdfFormFields.SetField(objBookMark.Descrizione, objBookMark.Valore)
                        pdfFormFields.SetField(objBookMark.Descrizione + "b", objBookMark.Valore)
                    Else
                        log.Debug("PrintPDF.splitto->" + objBookMark.Valore + " per ,")
                        Array = objBookMark.Valore.ToString().Split(", ")
                        Dim nOccur As Integer = 1
                        log.Debug("PrintPDF.devo prendere noccur da descrizione.IndexOf('|')->" + objBookMark.Descrizione)
                        If objBookMark.Descrizione.IndexOf("|") > 0 Then
                            nOccur = objBookMark.Descrizione.Substring(objBookMark.Descrizione.IndexOf("|") - 1, 1)
                        End If
                        log.Debug("PrintPDF.ho preso->" + nOccur.ToString)
                        pdfFormFields.SetField("T_DEBITI" + nOccur.ToString() + "|R", Array(0))
                        pdfFormFields.SetField("T_DEBITI" + nOccur.ToString() + "|Rb", Array(0))
                        pdfFormFields.SetField("T_DEBITIDEC" + nOccur.ToString() + "_SD", Array(1))
                        pdfFormFields.SetField("T_DEBITIDEC" + nOccur.ToString() + "_SDb", Array(1))
                        pdfFormFields.SetField("T_CODTRIBUTO" + nOccur.ToString(), objBookMark.CodTributo)
                        pdfFormFields.SetField("T_CODTRIBUTO" + nOccur.ToString() + "b", objBookMark.CodTributo)
                        '*** 20130422 - aggiornamento IMU ***
                        pdfFormFields.SetField("T_ANNORIF" + nOccur.ToString(), objBookMark.Anno)
                        pdfFormFields.SetField("T_ANNORIF" + nOccur.ToString() + "b", objBookMark.Anno)
                        pdfFormFields.SetField("T_NUMIMM" + nOccur.ToString(), objBookMark.NumFabb)
                        pdfFormFields.SetField("T_NUMIMM" + nOccur.ToString() + "b", objBookMark.NumFabb)
                        '*** ***
                        pdfFormFields.SetField("T_CODENTE" + nOccur.ToString() + "_SS", objBookMark.Ente)
                        pdfFormFields.SetField("T_CODENTE" + nOccur.ToString() + "_SSb", objBookMark.Ente)
                        pdfFormFields.SetField("T_SEZ" + nOccur.ToString() + "_SP", "EL")
                        pdfFormFields.SetField("T_SEZ" + nOccur.ToString() + "_SPb", "EL")
                        pdfFormFields.SetField("T_RATEAZ" + nOccur.ToString(), objBookMark.Rateizzazione)
                        pdfFormFields.SetField("T_RATEAZ" + nOccur.ToString() + "b", objBookMark.Rateizzazione)
                        pdfFormFields.SetField("T_ACC" + nOccur.ToString(), objBookMark.IsAcconto)
                        pdfFormFields.SetField("T_ACC" + nOccur.ToString() + "b", objBookMark.IsAcconto)
                        pdfFormFields.SetField("T_SAL" + nOccur.ToString(), objBookMark.IsSaldo)
                        pdfFormFields.SetField("T_SAL" + nOccur.ToString() + "b", objBookMark.IsSaldo)
                        pdfFormFields.SetField("T_RAVV" + nOccur.ToString(), objBookMark.isravvedimento)
                        pdfFormFields.SetField("T_RAVV" + nOccur.ToString() + "b", objBookMark.isravvedimento)
                    End If
                Else
                    log.Debug("objBookMark.Descrizione, objBookMark.Valore")
                    pdfFormFields.SetField(objBookMark.Descrizione, objBookMark.Valore)
                    pdfFormFields.SetField(objBookMark.Descrizione + "b", objBookMark.Valore)
                End If
            Next
            'flatten the form to remove editting options, set it to false to leave the form open to subsequent manual edits
            PDFStamper.FormFlattening = True
            'close the pdf
            PDFStamper.Close()

            oURLRet.Name = FileNameDOC
            oURLRet.Path = PathNameDOC
            oURLRet.Url = PathNameWEB
            oURLRet.oSetupDocumento = oggetto.TestataDOT.oSetupDocumento
            Return oURLRet
        Catch ex As Exception
            log.Debug("PrintPDF:: " & ex.Message, ex)
            Return Nothing
        End Try
    End Function
    'Private Function PrintPDF(ByVal oggetto As oggettoDaStampareCompleto, ByVal bIsStampaBollettino As Boolean, ByRef ArrModelli As objListModelliEsternalizza) As oggettoURL
    '    Try
    '        Dim oURLRet As New oggettoURL

    '        Dim AttoDOT As String = oggetto.TestataDOT.Atto
    '        Dim DominioDOT As String = oggetto.TestataDOT.Dominio
    '        Dim EnteDOT As String = oggetto.TestataDOT.Ente
    '        Dim FileNameDOT As String = oggetto.TestataDOT.Filename

    '        'per la gestione del file DOC
    '        Dim AttoDOC As String = oggetto.TestataDOC.Atto
    '        Dim DominioDOC As String = oggetto.TestataDOC.Dominio
    '        Dim EnteDOC As String = oggetto.TestataDOC.Ente
    '        Dim FileNameDOC As String = oggetto.TestataDOC.Filename

    '        Dim DataOra As String = DateTime.Now.ToString("ddMMyyyyHHmmss")

    '        'COMPONGO IL NOME DEL FILE TEMPLATE DA PRENDERE PER GENERARE IL DOCUMENTO
    '        Dim PathNameTemplate As String = ConstSession.CopyDir
    '        If AttoDOT.CompareTo("") <> 0 Then
    '            PathNameTemplate += AttoDOT + "\"
    '        End If
    '        If DominioDOT.CompareTo("") <> 0 Then
    '            PathNameTemplate += DominioDOT + "\"
    '        End If
    '        If EnteDOT.CompareTo("") <> 0 Then
    '            PathNameTemplate += EnteDOT + "\"
    '        End If
    '        If FileNameDOT.CompareTo("") <> 0 Then
    '            PathNameTemplate += FileNameDOT
    '        End If

    '        log.Debug("PathFileTemplate " & PathNameTemplate)

    '        'COMPONGO IL NOME DEL FILE DOC DA GENERARE E IL PERCORSO WEB
    '        Dim PathFileDOC As String = ConstSession.CopyDir
    '        Dim PathFileWEB As String = ConstSession.ExtDir
    '        Dim PathNameDOC As String = ""
    '        Dim PathNameWEB As String = ""
    '        If AttoDOC.CompareTo("") <> 0 Then
    '            PathFileDOC += AttoDOC + "\"
    '            PathFileWEB += AttoDOC + "/"
    '        End If
    '        If DominioDOC.CompareTo("") <> 0 Then
    '            PathFileDOC += DominioDOC + "\"
    '            PathFileWEB += DominioDOC + "/"
    '        End If
    '        If EnteDOC.CompareTo("") <> 0 Then
    '            PathFileDOC += EnteDOC + "\"
    '            PathFileWEB += EnteDOC + "/"
    '        End If
    '        CreateDir(PathFileDOC)
    '        log.Debug("Metto il ticks")
    '        'PathFileDOT = PathFileDOC
    '        FileNameDOC = FileNameDOC + DataOra + DateTime.Now.Ticks.ToString() + ConstSession.ManagedExtensions.PDF

    '        If FileNameDOC.CompareTo("") <> 0 Then
    '            PathNameDOC = PathFileDOC + FileNameDOC
    '            PathNameWEB = PathFileWEB + FileNameDOC
    '        End If

    '        log.Debug("PathFileDOC " & PathNameDOC)
    '        log.Debug("PathFileWEB " & PathNameWEB)

    '        log.Debug("Template copiato")
    '        log.Debug("PathFileDOC=" & PathFileDOC)
    '        log.Debug("PathFileTemplate=" & PathNameTemplate)
    '        log.Debug("FileNameDOC=" & FileNameDOC)
    '        log.Debug("PathFileWEB=" & PathFileWEB)

    '        'COPIO IL TEMPLATE NEL PERCORSO DI DESTINAZIONE
    '        Dim PDFStamper As New PdfStamper(New PdfReader(PathNameTemplate), New FileStream(PathNameDOC, FileMode.Create))

    '        Dim pdfFormFields As AcroFields
    '        Dim Array() As String
    '        Dim Scope As String
    '        Scope = ""
    '        For Each objBookMark As oggettiStampa In oggetto.Stampa
    '            pdfFormFields = PDFStamper.AcroFields
    '            log.Debug("pdfFormFields")
    '            If (objBookMark.Appartenenza <> "") Then
    '                log.Debug("PrintPDF.ho appartenenza->" + objBookMark.Appartenenza + " ho valore->" + objBookMark.Valore)
    '                Array = objBookMark.Valore.ToString().Split(",")
    '                If objBookMark.Descrizione.IndexOf("T_DEBITI") >= 0 Then
    '                    Scope = objBookMark.Appartenenza
    '                End If
    '                If (objBookMark.CodTributo = "dovutaTotale") Then
    '                    If Scope = objBookMark.Appartenenza Then
    '                        log.Debug("PrintPDF.ho scope=appartenenza->" + Scope + " splitto->" + objBookMark.Valore + " per ,")
    '                        Array = objBookMark.Valore.ToString().Split(", ")
    '                        pdfFormFields.SetField(objBookMark.Descrizione, Array(0))
    '                        pdfFormFields.SetField(objBookMark.Descrizione + "b", Array(0))
    '                        pdfFormFields.SetField("T_TOTSALDODEC_SD", Array(1))
    '                        pdfFormFields.SetField("T_TOTSALDODEC_SDb", Array(1))
    '                    End If
    '                ElseIf (objBookMark.CodTributo = "detrazione" Or objBookMark.Descrizione.IndexOf("IDOPERAZ") > 0) Then
    '                    pdfFormFields.SetField(objBookMark.Descrizione, objBookMark.Valore)
    '                    pdfFormFields.SetField(objBookMark.Descrizione + "b", objBookMark.Valore)
    '                Else
    '                    log.Debug("PrintPDF.splitto->" + objBookMark.Valore + " per ,")
    '                    Array = objBookMark.Valore.ToString().Split(", ")
    '                    Dim nOccur As Integer = 1
    '                    log.Debug("PrintPDF.devo prendere noccur da descrizione.IndexOf('|')->" + objBookMark.Descrizione)
    '                    If objBookMark.Descrizione.IndexOf("|") > 0 Then
    '                        nOccur = objBookMark.Descrizione.Substring(objBookMark.Descrizione.IndexOf("|") - 1, 1)
    '                    End If
    '                    log.Debug("PrintPDF.ho preso->" + nOccur.ToString)
    '                    pdfFormFields.SetField("T_DEBITI" + nOccur.ToString() + "|R", Array(0))
    '                    pdfFormFields.SetField("T_DEBITI" + nOccur.ToString() + "|Rb", Array(0))
    '                    pdfFormFields.SetField("T_DEBITIDEC" + nOccur.ToString() + "_SD", Array(1))
    '                    pdfFormFields.SetField("T_DEBITIDEC" + nOccur.ToString() + "_SDb", Array(1))
    '                    pdfFormFields.SetField("T_CODTRIBUTO" + nOccur.ToString(), objBookMark.CodTributo)
    '                    pdfFormFields.SetField("T_CODTRIBUTO" + nOccur.ToString() + "b", objBookMark.CodTributo)
    '                    '*** 20130422 - aggiornamento IMU ***
    '                    pdfFormFields.SetField("T_ANNORIF" + nOccur.ToString(), objBookMark.Anno)
    '                    pdfFormFields.SetField("T_ANNORIF" + nOccur.ToString() + "b", objBookMark.Anno)
    '                    pdfFormFields.SetField("T_NUMIMM" + nOccur.ToString(), objBookMark.NumFabb)
    '                    pdfFormFields.SetField("T_NUMIMM" + nOccur.ToString() + "b", objBookMark.NumFabb)
    '                    '*** ***
    '                    pdfFormFields.SetField("T_CODENTE" + nOccur.ToString() + "_SS", objBookMark.Ente)
    '                    pdfFormFields.SetField("T_CODENTE" + nOccur.ToString() + "_SSb", objBookMark.Ente)
    '                    pdfFormFields.SetField("T_SEZ" + nOccur.ToString() + "_SP", "EL")
    '                    pdfFormFields.SetField("T_SEZ" + nOccur.ToString() + "_SPb", "EL")
    '                    pdfFormFields.SetField("T_RATEAZ" + nOccur.ToString(), objBookMark.Rateizzazione)
    '                    pdfFormFields.SetField("T_RATEAZ" + nOccur.ToString() + "b", objBookMark.Rateizzazione)
    '                    pdfFormFields.SetField("T_ACC" + nOccur.ToString(), objBookMark.IsAcconto)
    '                    pdfFormFields.SetField("T_ACC" + nOccur.ToString() + "b", objBookMark.IsAcconto)
    '                    pdfFormFields.SetField("T_SAL" + nOccur.ToString(), objBookMark.IsSaldo)
    '                    pdfFormFields.SetField("T_SAL" + nOccur.ToString() + "b", objBookMark.IsSaldo)
    '                End If
    '            Else
    '                log.Debug("objBookMark.Descrizione, objBookMark.Valore")
    '                pdfFormFields.SetField(objBookMark.Descrizione, objBookMark.Valore)
    '                pdfFormFields.SetField(objBookMark.Descrizione + "b", objBookMark.Valore)
    '            End If
    '        Next
    '        'flatten the form to remove editting options, set it to false to leave the form open to subsequent manual edits
    '        PDFStamper.FormFlattening = True
    '        'close the pdf
    '        PDFStamper.Close()

    '        oURLRet.Name = FileNameDOC
    '        oURLRet.Path = PathNameDOC
    '        oURLRet.Url = PathNameWEB
    '        oURLRet.oSetupDocumento = oggetto.TestataDOT.oSetupDocumento
    '        Return oURLRet
    '    Catch ex As Exception
    '        log.Debug("PrintPDF:: " & ex.Message, ex)
    '        Return Nothing
    '    End Try
    'End Function
    ''' <summary>
    ''' converto il file in pdf e faccio il merge di tutti i file
    ''' </summary>
    ''' <param name="PathFileTestataDOC"></param>
    ''' <param name="oDocDaUnire"></param>
    ''' <returns></returns>
    Private Function UnionPDF(ByVal PathFileTestataDOC As String, ByVal oDocDaUnire As oggettoURL()) As Boolean
        'converto il file in pdf e faccio il merge di tutti i file
        Try
            Dim NamePDF As String
            Dim PathPDF As String
            Dim reader As PdfReader = Nothing
            Dim pageCount As Integer = 0
            Dim currentPage As Integer = 0
            Dim pdfDoc As iTextSharp.text.Document = Nothing
            Dim writer As PdfCopy = Nothing
            Dim page As PdfImportedPage = Nothing
            Dim currentPDF As Integer = 0

            If oDocDaUnire.Length > 0 Then
                PathPDF = Path.GetDirectoryName(oDocDaUnire(currentPDF).Path)
                If (Path.GetExtension(oDocDaUnire(currentPDF).Path).IndexOf(ConstSession.ManagedExtensions.PDF) < 0) Then
                    NamePDF = Path.GetFileNameWithoutExtension(oDocDaUnire(currentPDF).Path) + ConstSession.ManagedExtensions.PDF
                    log.Debug("devo convertire::" & oDocDaUnire(currentPDF).Path & " ::in::" & PathPDF + "\" + NamePDF)
                    EvenPage(oDocDaUnire(currentPDF).Path, PathPDF + "\" + NamePDF)
                Else
                    NamePDF = oDocDaUnire(currentPDF).Name
                End If

                reader = New iTextSharp.text.pdf.PdfReader(PathPDF + "\" + NamePDF)
                pdfDoc = New iTextSharp.text.Document(reader.GetPageSizeWithRotation(1))
                writer = New iTextSharp.text.pdf.PdfCopy(pdfDoc, New IO.FileStream(PathFileTestataDOC, IO.FileMode.OpenOrCreate, IO.FileAccess.Write))

                pageCount = reader.NumberOfPages

                While currentPDF < oDocDaUnire.Length
                    pdfDoc.Open()

                    While currentPage < pageCount
                        currentPage += 1
                        pdfDoc.SetPageSize(reader.GetPageSizeWithRotation(currentPage))
                        pdfDoc.NewPage()
                        page = writer.GetImportedPage(reader, currentPage)
                        writer.AddPage(page)
                    End While

                    currentPDF += 1
                    If currentPDF < oDocDaUnire.Length Then
                        PathPDF = Path.GetDirectoryName(oDocDaUnire(currentPDF).Path)
                        If (Path.GetExtension(oDocDaUnire(currentPDF).Path).IndexOf(ConstSession.ManagedExtensions.PDF) < 0) Then
                            NamePDF = Path.GetFileNameWithoutExtension(oDocDaUnire(currentPDF).Path) + ConstSession.ManagedExtensions.PDF
                            log.Debug("devo convertire::" & oDocDaUnire(currentPDF).Path & " ::in::" & PathPDF + "\" + NamePDF)
                            EvenPage(oDocDaUnire(currentPDF).Path, PathPDF + "\" + NamePDF)
                        Else
                            NamePDF = oDocDaUnire(currentPDF).Name
                        End If

                        reader = New iTextSharp.text.pdf.PdfReader(PathPDF + "\" + NamePDF)
                        pageCount = reader.NumberOfPages
                        currentPage = 0
                    End If
                End While

                pdfDoc.Close()
            End If
            log.Debug("salvo  in::" & PathFileTestataDOC)
            Return True
        Catch ex As Exception
            log.Debug("UnionPDF::si è verificato il seguente errore::" & ex.Message)
            Return False
        End Try
    End Function
#End Region
End Class
