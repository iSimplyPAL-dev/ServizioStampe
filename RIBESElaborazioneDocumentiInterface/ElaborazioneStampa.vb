Imports System
Imports System.Runtime.Remoting
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Http
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Collections
Imports System.Collections.ArrayList
Imports RIBESElaborazioneDocumentiInterface.Stampa.oggetti


''' <summary>
''' Definizione interfacce per la produzione dei documenti
''' </summary>
Public Interface IElaborazioneStampaDocOggetti
    '*** 201511 - template documenti per ruolo ***
    Function StampaDocumenti(ByVal PathTemplate As String, ByVal PathVirtualTemplate As String, ByVal TestataGruppo As oggettoTestata, ByVal GruppiDocumenti As GruppoDocumenti(), ByVal bIsStampaBollettino As Boolean, ByVal bCreaDPF As Boolean) As GruppoURL
    '*** ***
End Interface
