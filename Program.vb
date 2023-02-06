Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports NTSInformatica
Imports NTSInformatica.CLN__STD

Module Program

    Private oApp As CLE__APP = Nothing
    Private oMenu As CLE__MENU = Nothing
    Private documento As CLEVEBOLL

    Private _busConfig As String = String.Empty
    Private _dbBusiness As BusinessDbDataContext = Nothing
    Private _strSqlBusiness As String = String.Empty
    Private _nomeLogFile As String = String.Empty
    Private _pathLog As String = System.AppDomain.CurrentDomain.BaseDirectory & "LOGS"
    Private _dittaCorrente As String = String.Empty

    Sub Main()
        Try

            _strSqlBusiness = My.Settings.FANTICConnectionString

            CreaLogFile()
            PulisciFilesVecchi()

            If Not CollegaBusiness() Then
                wl("Impossibile connettersi al software gestionale BUSINESS. Programma interrotto")
                Return
            End If

            If Not TestCollegaSqlBusiness() Then
                wl("Impossibile connettersi al data base. Programma interrotto")
                Return
            End If

            ModificaDocumenti()


        Catch ex As Exception
            MessageBox.Show("ATTENZIONE!!! Errore irreversibile programma bloccato.", "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

#Region "METODI GESTIONE BUSINESS"

    Private Function TestCollegaSqlBusiness() As Boolean
        Try

            Dim dbBusiness As BusinessDbDataContext = New BusinessDbDataContext(_strSqlBusiness)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(TestCollegaSqlBusiness)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function CollegaBusiness() As Boolean

        Try

            'Collegamento a Business
            Dim strDateFormat As String = CurrentThread.CurrentCulture.DateTimeFormat.ShortDatePattern
            If strDateFormat.IndexOf("yyyy") = -1 Then
                strDateFormat = strDateFormat.Replace("yy", "yyyy")
                Dim oCulture As CultureInfo = New CultureInfo(CurrentThread.CurrentCulture.Name)
                oCulture.DateTimeFormat.ShortDatePattern = strDateFormat
                CurrentThread.CurrentCulture = oCulture
            End If

            Dim bOk As Boolean = False
            _busConfig = My.Settings.BusConfig
            _dittaCorrente = My.Settings.CodDitta

            'avvio il menu di Business
            Dim strError As String = ""
            oMenu = New CLE__MENU
            AddHandler oMenu.RemoteEvent, AddressOf GestisciEventiEntityAppBase
            If Not oMenu.Init(_busConfig, oApp, Nothing, strError) Then
                wl("Errore inizializzazione menu")
                wl("Dettaglio errore" & strError)
                Return False
            End If

            If Not oMenu.InitEx(strError, Nothing) Then
                wl("Errore inizializzazione menuExt")
                wl("Dettaglio errore" & strError)
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(CollegaBusiness)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub GestisciEventiEntityAppBase(sender As Object, ByRef e As NTSEventArgs)
        Try
            If e.Message <> String.Empty Then
                wl(String.Format("Errore interno App. Err: {0}", e.Message))
            End If
        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(GestisciEventiEntityAppBase)} . Errore: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "METODI LOGS"

    Private Sub PulisciFilesVecchi()
        Try
            Dim dtCreated As DateTime
            Dim dtToday As DateTime = Today.Date
            Dim flObj As FileInfo
            Dim ts As TimeSpan
            Dim lstDirsToDelete As New List(Of String)

            If Not Directory.Exists(_pathLog) Then
                Directory.CreateDirectory(_pathLog)
            End If

            For Each file As String In Directory.GetFiles(_pathLog, "*.txt")
                flObj = New FileInfo(file)
                dtCreated = flObj.CreationTime

                ts = dtToday - dtCreated

                'Add whatever storing you want here for all folders...

                If ts.Days > 120 Then
                    wl("Cancellazione File di log : " & file)
                    flObj.Delete()
                End If
            Next

        Catch ex As Exception
            MessageBox.Show("Errore creazione file di log. Errore: " & ex.Message, "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub wl(ByVal line As String)
        Try
            'Dim f As File = New File
            Dim line1 As String = DateTime.Now.ToString() & " ---- " & line & Environment.NewLine
            Using writer As System.IO.StreamWriter = File.AppendText(_nomeLogFile)
                writer.WriteLine(line1)
            End Using
        Catch ex As Exception
            MessageBox.Show("Errore scrittura file di log. Errore: " & ex.Message, "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub CreaLogFile()
        Dim nomeFile As String = _pathLog & "\Utility_Ordini_Mod_Prezzi_Log" & DateTime.Now.Day.ToString & "_" & DateTime.Now.Month.ToString & "_" & DateTime.Now.Year & "_" & DateTime.Now.Hour.ToString & "_" & DateTime.Now.Minute.ToString & "_" & DateTime.Now.Second.ToString & ".txt"
        _nomeLogFile = nomeFile
        Try
            Dim f As FileStream = File.Create(_nomeLogFile)
            f.Close()
        Catch ex As Exception
            MessageBox.Show("ATTENZIONE!!! Impossibile creare file di log. Errore: " & ex.Message, "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

#End Region

#Region "ELABORAZIONE DOCUMENTI"

    Private Function ModificaDocumenti() As Boolean
        Try

            Dim listaDocumenti As List(Of RDM_Param9100) = Nothing

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)
                dbBusiness.CommandTimeout = 30000

                listaDocumenti = dbBusiness.RDM_Param9100.Where(Function(x) x.ValoreFatEuro = 0 And
                                                                            x.SerieDoc = "B" And
                                                                            x.NumDoc = 4151).ToList()
            End Using

            ElaboraDocumenti(listaDocumenti)

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(ElaboraDocumenti)}. Errore: {ex.Message}")
            Return False
        End Try
    End Function

    Public Function ElaboraDocumenti(ByVal listaDoc As List(Of RDM_Param9100)) As Boolean
        Try
            Dim bResultOk As Boolean = True

            For Each doc As RDM_Param9100 In listaDoc

                initDocumento()

                Dim dsDoc As New DataSet
                If Not documento.ApriDoc(oApp.Ditta, False, "T", doc.AnnoDoc, doc.SerieDoc, doc.NumDoc, dsDoc) Then
                    bResultOk = False
                    wl("Non è stato possibile aprire il documento tipo:T, anno:{doc.AnnoDoc}, serie:{doc.SerieDoc}, numero:{doc.NumDoc}.")
                    Continue For
                End If

                If Not ElaboraRigheCorpoDocumento() Then
                    bResultOk = False
                    wl("Errore, non tutte le righe del copro del documento tipo:T, anno:{doc.AnnoDoc}, serie:{doc.SerieDoc}, numero:{doc.NumDoc} sono state elaborate.")
                End If

                documento.CalcolaTotali()
                If Not documento.SalvaDocumento("U") Then
                    wl("Errore durante il salvataggio! Documento tipo:T, anno:{doc.AnnoDoc}, serie:{doc.SerieDoc}, numero:{doc.NumDoc}, non salvato.")
                    bResultOk = False
                End If

                documento.Dispose()
                documento = Nothing

            Next

            Return bResultOk

        Catch ex As Exception

            CLN__STD.GestErr(ex, "", "")

        End Try
    End Function

    Public Function ElaboraRigheCorpoDocumento() As Boolean
        Try
            Dim bResult As Boolean = True

            'TODOUTILITYMARCO
            documento.bDocNonModificabile = False
            documento.bInApriDocSilent = True

            Dim dttArticolo As New DataTable
            Dim dttConto As New DataTable

            For Each rigaCorpo As DataRow In documento.dttEC.Rows
                ''If NTSCInt(rigaCorpo!ec_provv) = 0 Then

                If Not oMenu.ValCodiceDb(NTSCStr(rigaCorpo!ec_codart), _dittaCorrente, "ARTICO", "S", "", dttArticolo) Then
                    bResult = False
                    Continue For
                End If

                If NTSCStr(dttArticolo.Rows(0)!ar_gesubic) = "N" Then
                    If NTSCStr(rigaCorpo!ec_ubicaz).Trim() <> String.Empty Then
                        rigaCorpo!ec_ubicaz = "  "
                    End If
                End If

                '''End If
            Next

            Return bResult

        Catch ex As Exception

            CLN__STD.GestErr(ex, "", "")

        End Try
    End Function

    Private Function initDocumento() As Boolean
        Try
            If Not documento Is Nothing Then Return True

            documento = New CLEVEBOLL
            AddHandler documento.RemoteEvent, AddressOf GestisciEventiEntityAppBase

            If Not documento.Init(oApp, Nothing, oMenu.oCleComm, "", False, "", "") Then Return False
            If Not documento.InitExt() Then Return False
            documento.bModuloCRM = False
            documento.bIsCRMUser = False
            oMenu = CType(oApp.oMenu, CLE__MENU)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(initDocumento)}. Errore: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

    Public Class DocumentoX

        Public Property Tipork As String
        Public Property Serie As String
        Public Property Anno As Integer
        Public Property Numero As Integer

    End Class

End Module
