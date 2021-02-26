Imports System.IO
Imports System.Security
Namespace Acao.Code.Util

    Public Class UtilArquivos

        Private Sub SubSelecionarPasta()
            'Me.ofd1.Multiselect = True
            ' Me.ofd1.Title = "Selecionar Pasta"
            'ofd1.InitialDirectory = "C:\"
            Dim ofd1 As New OpenFileDialog With {
                .Filter = "Texts (*.*)|*.txt ",
                .ValidateNames = False,
                .CheckFileExists = False,
                .FileName = "Selecionar Pasta.",
                .CheckPathExists = True
            }

            'ofd1.FilterIndex = 2
            'ofd1.RestoreDirectory = True
            'ofd1.ReadOnlyChecked = True
            'ofd1.ShowReadOnly = True
            Dim dr As DialogResult = ofd1.ShowDialog()
            If dr = System.Windows.Forms.DialogResult.OK Then
                Try
                    Replace(ofd1.FileName, "Selecionar Pasta.. ", "")
                    ofd1.Dispose()
                    ' Aqui fica o que deve ser executado com os arquivos selecionados.
                Catch ex As SecurityException
                    ' O usuário  não possui permissão para ler arquivos
                    MessageBox.Show((("Erro de segurança Contate o administrador de segurança da rede." & vbLf & vbLf & "Mensagem : ") + ex.Message & vbLf & vbLf & "Detalhes (enviar ao suporte):" & vbLf & vbLf) + ex.StackTrace)
                Catch ex As Exception
                    ' Não pode carregar o arquivo (problemas de permissão)
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End Sub
        Public Function SelecionarTXT() As String

            'Me.ofd1.Multiselect = True
            ' Me.ofd1.Title = "Selecionar Pasta"
            'ofd1.InitialDirectory = "C:\"
            Dim strFilePath As String = ""
            Dim ofd1 As New OpenFileDialog With {
                .Filter = "Text (*.txt,*csv)|*.txt;*.csv",
                .ValidateNames = False,
                .CheckFileExists = True,
                .CheckPathExists = True,
                .Multiselect = False
            }

            Dim dr As DialogResult = ofd1.ShowDialog()
            If dr = System.Windows.Forms.DialogResult.OK Then
                Try
                    strFilePath = ofd1.FileName.ToString
                    ofd1.Dispose()
                    ' Aqui fica o que deve ser executado com os arquivos selecionados.
                Catch ex As SecurityException
                    ' O usuário  não possui permissão para ler arquivos
                    MessageBox.Show((("Erro de segurança Contate o administrador de segurança da rede." & vbLf & vbLf & "Mensagem : ") + ex.Message & vbLf & vbLf & "Detalhes (enviar ao suporte):" & vbLf & vbLf) + ex.StackTrace)
                Catch ex As Exception
                    ' Não pode carregar o arquivo (problemas de permissão)
                    MessageBox.Show(ex.Message)
                End Try
            End If
            Return strFilePath
        End Function

        Public Function SelecionarPasta() As String
            Dim strPastas As String = ""
            Dim ofd1 As New OpenFileDialog With {
                .Filter = "Pasta (*.*)|*. ",
                .ValidateNames = False,
                .CheckFileExists = False,
                .FileName = "Selecionar Pasta.",
                .CheckPathExists = True,
                .ShowReadOnly = True
            }
            Dim dr As DialogResult = ofd1.ShowDialog()
            If dr = System.Windows.Forms.DialogResult.OK Then
                Try
                    'strPastas = Replace(ofd1.FileName, "Selecionar Pasta.. ", "")
                    strPastas = ofd1.FileName
                    ofd1.Dispose()
                Catch ex As SecurityException
                    Return ""
                    ' O usuário  não possui permissão para ler arquivos
                    Throw New Exception((("Erro de segurança Contate o administrador de segurança da rede." & vbLf & vbLf & "Mensagem : ") + ex.Message & vbLf & vbLf & "Detalhes (enviar ao suporte):" & vbLf & vbLf) + ex.StackTrace)
                Catch ex As Exception
                    Return ""
                    ' Não pode carregar o arquivo (problemas de permissão)
                    Throw New Exception(ex.Message)
                End Try
            End If
            Return Replace(strPastas, "\Selecionar Pasta", "")
        End Function

        Public Function ImportarFile(FolderPath As String, FiltraExtensao As String) As List(Of String)
            Try
                Dim files = Directory.EnumerateFiles(FolderPath, FiltraExtensao, SearchOption.AllDirectories).ToList
                Return files
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Erro ao ImportarXML ", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End Try
        End Function

    End Class
End Namespace
