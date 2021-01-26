Option Strict Off
Option Explicit On

Friend Class frmImportadorAutomatico
	Inherits System.Windows.Forms.Form
	Dim db As New ADODB.Connection
	Dim dbIntegrador As New ADODB.Recordset
	Dim IconeAtual As Short

    Public Sub Verifica(ByVal Horas As Short, ByVal Posicao As Double)
        Dim CaminhoAdo As Object
        Dim dbSql As New ADODB.Connection
        Dim db As New ADODB.Connection
        Dim dbFechamentos As New ADODB.Recordset
        Dim dbCaixas As New ADODB.Recordset
        Dim dbPdvs As New ADODB.Recordset
        Dim dbTurnos As New ADODB.Recordset
        Dim CodigoPosto As String

        Dim DataCaixa As Date
        Dim Turno As String
        Dim ProcimaImportacao As Date
        Dim CodigoTurno As Double
        ProcimaImportacao = DateAdd(Microsoft.VisualBasic.DateInterval.Hour, Horas, Now)

        If Me.ImportarBindingSource.Count > 0 Then
            For Posicao = 0 To Me.ImportarBindingSource.Count - 1
                With Me.ImportacaoDataSet.Importar(Posicao)
                    If Posicao >= 0 Then
                        If .IsUltimaImportacaoNull = True Then
                            .UltimaImportacao = CDate("01/01/2009 00:00:00")
                            Me.TableAdapterManager.UpdateAll(Me.ImportacaoDataSet)
                        End If
                        If .UltimaImportacao <= ProcimaImportacao Then
                            If Trim(.LocalDB) <> "" Then
                                CaminhoAdo = strMdb & .LocalDB
                                Try
                                    db.Close()
                                Catch ex As Exception

                                End Try

                                db.Open(CaminhoAdo)

                                dbPdvs.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                dbPdvs.Open("select *from pdvs", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                                dbTurnos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                dbTurnos.Open("select *from turnos order by horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                                dbFechamentos.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                dbFechamentos.Open("select *from fechamentodecaixa order by datacaixa, horaini", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                                If dbFechamentos.RecordCount = 0 Then
                                    If dbTurnos.RecordCount <> 0 Then
                                        dbTurnos.MoveFirst()
                                        DataCaixa = CDate("01/01/2009")
                                        Turno = dbTurnos.Fields("Descri").Value
                                        CodigoTurno = dbTurnos.Fields("CodigoTurno").Value
                                        .UltimoCaixa = FormatDateTime(DataCaixa.ToString, DateFormat.ShortDate) & " - " & Turno
                                    Else
                                        DataCaixa = CDate("01/01/2009")
                                        Turno = "01"
                                        .UltimoCaixa = FormatDateTime(DataCaixa.ToString, DateFormat.ShortDate) & " - " & Turno
                                        CodigoTurno = 1
                                    End If
                                Else
                                    dbFechamentos.MoveLast()
                                    DataCaixa = dbFechamentos.Fields("DataCaixa").Value
                                    Turno = dbFechamentos.Fields("Turno").Value
                                    CodigoTurno = dbFechamentos.Fields("CodigoTurno").Value
                                    .UltimoCaixa = FormatDateTime(DataCaixa.ToString, DateFormat.ShortDate) & " - " & Turno
                                End If
                                dbFechamentos.Close()
                                dbFechamentos.Open("select *from config", db, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
                                If dbFechamentos.RecordCount <> 0 Then
                                    strSql = "Provider=SQLOLEDB.1;Password=masterkey;Persist Security Info=True;User ID=sa;Initial Catalog=Integrador;Data Source=" & dbFechamentos.Fields("ftp").Value
                                    CodigoPosto = dbFechamentos.Fields("porta").Value
                                Else
                                    CodigoPosto = 0
                                    strSql = ""
                                End If
                                dbFechamentos.Close()

                                If strSql <> "" Then
                                    dbSql.Open(strSql)
                                    dbCaixas.CursorLocation = ADODB.CursorLocationEnum.adUseClient
                                    dbCaixas.Open("select datacaixa, turno, planodeconta from caixas where codigoposto='" & CodigoPosto & "' and datacaixa>='" & DataCaixa & "' group by datacaixa, turno, planodeconta order by datacaixa, turno, planodeconta", dbSql, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)

                                    If dbCaixas.RecordCount <> 0 Then
                                        Do While dbCaixas.Fields("Turno").Value <= Turno And dbCaixas.Fields("DataCaixa").Value = DataCaixa
                                            dbCaixas.MoveNext()
                                            If dbCaixas.EOF = True Then Exit Do
                                        Loop
                                        Do While dbCaixas.EOF = False
                                            If dbCaixas.Fields("DataCaixa").Value = DataCaixa Then
                                                If dbCaixas.Fields("Turno").Value > Turno Then
                                                    dbPdvs.MoveFirst()
                                                    dbPdvs.Find("codigo='" & dbCaixas.Fields("planodeconta").Value & "'")
                                                    dbTurnos.MoveFirst()
                                                    dbTurnos.Find("descri='" & dbCaixas.Fields("Turno").Value & "'")
                                                    If dbPdvs.EOF = False And dbTurnos.EOF = False Then
                                                        Importar(CaminhoAdo, dbCaixas.Fields("DataCaixa").Value, dbPdvs.Fields("Codigo").Value, dbTurnos.Fields("Descri").Value, dbTurnos.Fields("CodigoTurno").Value)

                                                        Try
                                                        Catch ex As Exception

                                                        End Try

                                                    End If
                                                End If
                                            Else
                                                dbPdvs.MoveFirst()
                                                dbPdvs.Find("codigo='" & dbCaixas.Fields("planodeconta").Value & "'")
                                                dbTurnos.MoveFirst()
                                                dbTurnos.Find("descri='" & dbCaixas.Fields("Turno").Value & "'")
                                                If dbPdvs.EOF = False And dbTurnos.EOF = False Then
                                                    Importar(CaminhoAdo, dbCaixas.Fields("DataCaixa").Value, dbPdvs.Fields("Codigo").Value, dbTurnos.Fields("Descri").Value, dbTurnos.Fields("CodigoTurno").Value)

                                                    Try
                                                    Catch ex As Exception
                                                        MsgBox(ex.Message)
                                                    End Try

                                                End If
                                            End If
                                            dbCaixas.MoveNext()
                                        Loop
                                        dbPdvs.Close()
                                        db.Close()
                                    End If
                                    dbCaixas.Close()
                                    dbSql.Close()
                                End If

                            End If
                        End If
                        .UltimaImportacao = Now


                    End If

                End With
            Next
        End If
        Me.TableAdapterManager.UpdateAll(Me.ImportacaoDataSet)

        Try

            

        Catch ex As Exception

        End Try


    End Sub

	Private Sub cmdVefivicar_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdVefivicar.Click
		Dim Horas As Short
		Horas = 0
        Timer1.Enabled = False
        Verifica(Horas, 0)
        Try

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

		Timer1.Enabled = True
	End Sub
	
	Private Sub frmImportadorAutomatico_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Me.ImportarTableAdapter.Fill(Me.ImportacaoDataSet.Importar)

        Configura.ChequesNoCaixa = CShort(ReadINI("cheques", "Cheques", 0, My.Application.Info.DirectoryPath & "\Posto.ini"))
		Configura.NotaNoCaixa = CShort(ReadINI("Notas no Caixa", "Nocaixa", 0, My.Application.Info.DirectoryPath & "\Posto.ini"))
		Configura.NotaBloqueia = CShort(ReadINI("Notas no Caixa", "Bloqueia", 0, My.Application.Info.DirectoryPath & "\Posto.ini"))
		
		strMdb = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source="
		
		CaminhoImporta = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & My.Application.Info.DirectoryPath & "\Importacao.mdb"
		
		txtHoras.Text = GetSetting(My.Application.Info.AssemblyName, "Configura", "Horas", "3")
		
		Me.Hide()
	End Sub
	
	Private Sub frmImportadorAutomatico_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        If Me.WindowState = System.Windows.Forms.FormWindowState.Minimized Then Me.Hide()
    End Sub
	
	Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        Dim Horas As Short
        If IsNumeric(txtHoras.Text) = False Then
            Horas = -3
        Else
            Horas = -CShort(txtHoras.Text)
        End If
        Timer1.Enabled = False
        Try
            Verifica(Horas, 0)
        Catch ex As Exception

        End Try

        Timer1.Enabled = True
    End Sub
	
	Private Sub txtHoras_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtHoras.TextChanged
        SaveSetting(My.Application.Info.AssemblyName, "Configura", "Horas", txtHoras.Text)
    End Sub
	
    Private Sub ConfigurarToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ConfigurarToolStripMenuItem1.Click
        Me.Show()
        Me.Activate()
    End Sub

    Private Sub FecharOProgramaToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FecharOProgramaToolStripMenuItem1.Click
        Dim Resposta As Short
        Resposta = MsgBox("Deseja sair do programa!", MsgBoxStyle.YesNo)
        If Resposta = MsgBoxResult.Yes Then
            End
        End If
    End Sub

    Private Sub ImportarBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportarBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.ImportarBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.ImportacaoDataSet)

    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
    End Sub

End Class