Imports DanfeSharp
Imports DanfeSharp.Modelo
Imports GerarDANFE.Acao.Code.DAL
Imports GerarDANFE.Acao.Code.Util


Public Class Form1
    Dim Pasta As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim Util As New UtilArquivos
            Dim Files As List(Of String)
            Dim Validar As New ValidarNFe
            Dim NFeLista As New List(Of DanfeViewModel)
            Pasta = Util.SelecionarPasta & "\"
            Files = Util.ImportarFile(Pasta, "*.xml")

            If Files.Count > 0 Then
                For Each file As String In Files
                    Validar.OpenXML(file)
                    If Validar.NFeValidado Then
                        Dim Modelo = DanfeViewModelCreator.CriarDeArquivoXml(file)
                        NFeLista.Add(Modelo)
                    End If
                Next
            End If
            SalvarDanfe(NFeLista)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        MessageBox.Show("Fim")
    End Sub

    Sub SalvarDanfe(ByVal ListModelo As List(Of DanfeViewModel))
        Dim TotalFile As Integer = 0
        Dim ContFile As Integer = 0
        Dim Perc As Integer
        TotalFile = ListModelo.Count
        For Each ItemModelo As DanfeViewModel In ListModelo
            Using danfe = New Danfe(ItemModelo)
                danfe.Gerar()
                Dim stChave = danfe.ViewModel.ChaveAcesso
                danfe.Salvar(Pasta & "\" & stChave & ".Pdf")
            End Using
            ContFile += 1
            Perc = ((ContFile / TotalFile) * 100)
            If Perc Mod 10 = 0 Then
                pgb.Value = Perc
                pgb.Refresh()
                Application.DoEvents()
            End If
        Next
        ListModelo = Nothing
    End Sub
End Class
