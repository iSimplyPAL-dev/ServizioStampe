﻿'------------------------------------------------------------------------------
'<auto-generated>
'    Il codice è stato generato da uno strumento.
'    Versione runtime:4.0.30319.42000
'
'    Le modifiche apportate a questo file possono provocare un comportamento non corretto e andranno perse se
'    il codice viene rigenerato.
'</auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'Il codice sorgente è stato generato automaticamente da Microsoft.VSDesigner, versione 4.0.30319.42000.
'
Namespace WSPrintBarcode
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2558.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="ServiceStampaBarcodeSoap", [Namespace]:="http://tempuri.org/")>  _
    Partial Public Class ServiceStampaBarcode
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        Private PrintBarcodeOperationCompleted As System.Threading.SendOrPostCallback
        
        Private useDefaultCredentialsSetExplicitly As Boolean
        
        '''<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = "https://www.ran.it/wsstampabarcode/servicestampabarcode.asmx"
            If (Me.IsLocalFileSystemWebService(Me.Url) = true) Then
                Me.UseDefaultCredentials = true
                Me.useDefaultCredentialsSetExplicitly = false
            Else
                Me.useDefaultCredentialsSetExplicitly = true
            End If
        End Sub
        
        Public Shadows Property Url() As String
            Get
                Return MyBase.Url
            End Get
            Set
                If (((Me.IsLocalFileSystemWebService(MyBase.Url) = true)  _
                            AndAlso (Me.useDefaultCredentialsSetExplicitly = false))  _
                            AndAlso (Me.IsLocalFileSystemWebService(value) = false)) Then
                    MyBase.UseDefaultCredentials = false
                End If
                MyBase.Url = value
            End Set
        End Property
        
        Public Shadows Property UseDefaultCredentials() As Boolean
            Get
                Return MyBase.UseDefaultCredentials
            End Get
            Set
                MyBase.UseDefaultCredentials = value
                Me.useDefaultCredentialsSetExplicitly = true
            End Set
        End Property
        
        '''<remarks/>
        Public Event PrintBarcodeCompleted As PrintBarcodeCompletedEventHandler
        
        '''<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://tempuri.org/PrintBarcode", RequestNamespace:="http://tempuri.org/", ResponseNamespace:="http://tempuri.org/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function PrintBarcode(ByVal nTypeCode As Integer, ByVal sDaCodificare As String, ByVal sPathFile As String, ByVal sNameFile As String, ByVal sBookmark As String) As Boolean
            Dim results() As Object = Me.Invoke("PrintBarcode", New Object() {nTypeCode, sDaCodificare, sPathFile, sNameFile, sBookmark})
            Return CType(results(0),Boolean)
        End Function
        
        '''<remarks/>
        Public Function BeginPrintBarcode(ByVal nTypeCode As Integer, ByVal sDaCodificare As String, ByVal sPathFile As String, ByVal sNameFile As String, ByVal sBookmark As String, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("PrintBarcode", New Object() {nTypeCode, sDaCodificare, sPathFile, sNameFile, sBookmark}, callback, asyncState)
        End Function
        
        '''<remarks/>
        Public Function EndPrintBarcode(ByVal asyncResult As System.IAsyncResult) As Boolean
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),Boolean)
        End Function
        
        '''<remarks/>
        Public Overloads Sub PrintBarcodeAsync(ByVal nTypeCode As Integer, ByVal sDaCodificare As String, ByVal sPathFile As String, ByVal sNameFile As String, ByVal sBookmark As String)
            Me.PrintBarcodeAsync(nTypeCode, sDaCodificare, sPathFile, sNameFile, sBookmark, Nothing)
        End Sub
        
        '''<remarks/>
        Public Overloads Sub PrintBarcodeAsync(ByVal nTypeCode As Integer, ByVal sDaCodificare As String, ByVal sPathFile As String, ByVal sNameFile As String, ByVal sBookmark As String, ByVal userState As Object)
            If (Me.PrintBarcodeOperationCompleted Is Nothing) Then
                Me.PrintBarcodeOperationCompleted = AddressOf Me.OnPrintBarcodeOperationCompleted
            End If
            Me.InvokeAsync("PrintBarcode", New Object() {nTypeCode, sDaCodificare, sPathFile, sNameFile, sBookmark}, Me.PrintBarcodeOperationCompleted, userState)
        End Sub
        
        Private Sub OnPrintBarcodeOperationCompleted(ByVal arg As Object)
            If (Not (Me.PrintBarcodeCompletedEvent) Is Nothing) Then
                Dim invokeArgs As System.Web.Services.Protocols.InvokeCompletedEventArgs = CType(arg,System.Web.Services.Protocols.InvokeCompletedEventArgs)
                RaiseEvent PrintBarcodeCompleted(Me, New PrintBarcodeCompletedEventArgs(invokeArgs.Results, invokeArgs.Error, invokeArgs.Cancelled, invokeArgs.UserState))
            End If
        End Sub
        
        '''<remarks/>
        Public Shadows Sub CancelAsync(ByVal userState As Object)
            MyBase.CancelAsync(userState)
        End Sub
        
        Private Function IsLocalFileSystemWebService(ByVal url As String) As Boolean
            If ((url Is Nothing)  _
                        OrElse (url Is String.Empty)) Then
                Return false
            End If
            Dim wsUri As System.Uri = New System.Uri(url)
            If ((wsUri.Port >= 1024)  _
                        AndAlso (String.Compare(wsUri.Host, "localHost", System.StringComparison.OrdinalIgnoreCase) = 0)) Then
                Return true
            End If
            Return false
        End Function
    End Class
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2558.0")>  _
    Public Delegate Sub PrintBarcodeCompletedEventHandler(ByVal sender As Object, ByVal e As PrintBarcodeCompletedEventArgs)
    
    '''<remarks/>
    <System.CodeDom.Compiler.GeneratedCodeAttribute("System.Web.Services", "4.7.2558.0"),  _
     System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code")>  _
    Partial Public Class PrintBarcodeCompletedEventArgs
        Inherits System.ComponentModel.AsyncCompletedEventArgs
        
        Private results() As Object
        
        Friend Sub New(ByVal results() As Object, ByVal exception As System.Exception, ByVal cancelled As Boolean, ByVal userState As Object)
            MyBase.New(exception, cancelled, userState)
            Me.results = results
        End Sub
        
        '''<remarks/>
        Public ReadOnly Property Result() As Boolean
            Get
                Me.RaiseExceptionIfNecessary
                Return CType(Me.results(0),Boolean)
            End Get
        End Property
    End Class
End Namespace
