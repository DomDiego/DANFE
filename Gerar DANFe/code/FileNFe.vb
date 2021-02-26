Imports System.Text
Imports System.Xml

Namespace Acao.Code.DAL
    Public Class ValidarNFe
        Property NFeValidado As Boolean
        Private Function XmleNFE(doc As XmlDocument) As Boolean
            Dim eNFe As Boolean
            Try
                'NodNFe = Nothing
                eNFe = False
                For Each nodeRoot As XmlNode In doc.ChildNodes
                    'notas sem nfeProc so com NFe
                    If nodeRoot.LocalName.Equals("NFe") Then
                        'ProcessaNFe(nodeRoot)
                        ' NodNFe = nodeRoot
                        eNFe = True
                    End If
                    'notas com nó nfeProc /nfe / ChildNodes.Count > 1 para evitar erros ao ler xml de eventos
                    If nodeRoot.LocalName.Equals("nfeProc") And nodeRoot.ChildNodes.Count > 1 Then
                        ' ProcessaNFe(nodeRoot.ChildNodes.Item(0))
                        ' NodNFe = nodeRoot.ChildNodes.Item(0)
                        eNFe = True
                    End If
                    'notas com log
                    If nodeRoot.LocalName.Equals("retConsNFeLog") Then
                        'ler layout 
                        For Each noderetConsNFeLog As XmlNode In nodeRoot.ChildNodes
                            'nfelog
                            For Each nodeNFeLog As XmlNode In noderetConsNFeLog.ChildNodes
                                'nfeproc
                                If nodeNFeLog.LocalName.Equals("nfeProc") Then
                                    'ProcessaNFe(nodeNFeLog.ChildNodes.Item(0))
                                    ' NodNFe = nodeNFeLog.ChildNodes.Item(0)
                                    eNFe = True
                                End If
                                For Each noderetProc As XmlNode In nodeNFeLog.ChildNodes
                                    If noderetProc.LocalName.Equals("proc") Then
                                        For Each proc As XmlNode In noderetProc.ChildNodes
                                            'Count > 1 para pular xml de eventos 
                                            If proc.LocalName.Equals("nfeProc") And proc.ChildNodes.Count > 1 Then
                                                'ProcessaNFe(proc.ChildNodes.Item(0))
                                                'NodNFe = proc.ChildNodes.Item(0)
                                                eNFe = True
                                            End If
                                        Next
                                    End If
                                Next
                            Next
                        Next
                    End If
                Next
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
            Return eNFe
        End Function
        Public Sub OpenXML(pCaminho As String)
            Try
                Using ostream As New IO.StreamReader(pCaminho, Encoding.UTF8)
                    Dim doc = New XmlDocument()
                    doc.LoadXml(ostream.ReadToEnd.ToString)
                    NFeValidado = XmleNFE(doc)
                    doc = Nothing
                End Using
            Catch ex As Exception
            End Try
        End Sub
    End Class

End Namespace

