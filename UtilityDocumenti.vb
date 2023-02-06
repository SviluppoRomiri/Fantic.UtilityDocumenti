Imports System.Globalization
Imports System.IO
Imports System.Text
Imports System.Threading.Thread
Imports System.Windows.Forms
Imports NTSInformatica
Imports NTSInformatica.CLN__STD
Imports Renci.SshNet
Imports Renci.SshNet.Sftp
Imports System.Net.Mail
Imports FastMember

Public Class GestFanticDhl

#Region "ATTRIBUTI"

    Private oApp As CLE__APP = Nothing
    Private oMenu As CLE__MENU = Nothing
    Public docMagVeboll As CLEVEBOLL
    Private _busConfig As String = String.Empty
    Private _dbBusiness As BusinessDbDataContext = Nothing
    Private _strSqlBusiness As String = String.Empty
    Private _nomeLogFile As String = String.Empty
    Private _pathLog As String = System.AppDomain.CurrentDomain.BaseDirectory & "LOGS"
    Private _dittaCorrente As String = String.Empty
    Private _flagErroriOk As Boolean = True

    Public _pathLocaleCSV As String = String.Empty
    Public _pathInputRemotoCSV As String = String.Empty

    Public _fileDiRitorno As String = String.Empty

    Public _magazzinoDhl As Integer = 9990
    Public _magzzinoB2CDhl As Integer = 9980

    Private Const MOTO_CARICO As String = "MOTOCamion"
    Private Const MONO_CARICO As String = "MONOCamion"

    Private Const MOTO_SPED As String = "fmspedMB_"
    Private Const MONO_SPED As String = "fmspedB2B_"

    Private Const B2B_TRASF As String = "fmassign_"
    Private Const MERCATO_B2B As String = "B2B"
    Private Const MERCATO_B2C As String = "B2C"

    Private Const INTESTAZIONE_CAMION_CARICO As String = "RifKey|RifBolla|CodArt|Matricola"
    Private Const INTESTAZIONE_SPEDIZIONE As String = "RifKey|RagSociale|Indirizzo|Cap|Localita|StatoIso|NumColli|Peso|Volume|Note|IdMail|DataCons|Telefono|SpondaIdraulica|CodAccuntDhl|CodArt"
    Private Const INTESTAZIONE_TRASF_B2C As String = "CodArt|Matricola|Mercato"

    Private Const PATH_USCITA_MOTO_DA_FANTIC As String = "moto/in/"
    Private Const PATH_USCITA_MOTO_DA_DHL As String = "moto/out/"
    Private Const PATH_USCITA_MONO_DA_FANTIC As String = "monopattini/in/"
    Private Const PATH_USCITA_MONO_DA_DHL As String = "monopattini/out/"

    'Costanti BMI CARICO DHL
    Private Const BMI_TIPODOC As String = "Z"
    Private Const BMI_SERIE As String = "@"
    Private Const BMI_CONTO As Integer = 21069998
    Private Const BMI_TPBF As Integer = 910

#End Region

    Public Overridable Function GestioneFanticDhl() As Boolean
        Try

            CreaLogFile()
            PulisciFilesVecchi()

            _magazzinoDhl = My.Settings.MagDhlVr
            _magzzinoB2CDhl = My.Settings.MagB2CDhl
            _strSqlBusiness = My.Settings.FANTICConnectionString
            _pathLocaleCSV = My.Settings.PathCSV
            _pathInputRemotoCSV = My.Settings.FtpPath

            If Not CollegaBusiness() Then
                wl("Impossibile connettersi al software gestionale BUSINESS. Programma interrotto")
                Return False
            End If

            If Not TestCollegaSqlBusiness() Then
                wl("Impossibile connettersi al data base. Programma interrotto")
                Return False
            End If

            If Not MotoEsportaCsvCaricoMagDhlVerona() Then
                wl("Errore. Carico Moto magazzino di Verona terminato con errori.")
            End If

            If Not MonopattiniEsportaCsvCaricoMagDhlVerona() Then
                wl("Errore. Carico Monopattini magazzino di Verona terminato con errori.")
            End If

            If Not MotoEsportCsvUsciteDaDhlVerona() Then
                wl("Errore. Invio file uscite Moto terminato con errori.")
            End If

            If Not MonopattiniEsportCsvUsciteDaDhlVerona() Then
                wl("Errore. Invio file uscite Monopattini terminato con errori.")
            End If

            If Not MercatiDocCaricoB2bB2c() Then
                wl("Errore. Trasferimento monopattini da B2B a B2C terminato con errori.")
            End If

            'Pausa 1 minuto 
            Sleep(60000)
            'Leggo tutti i file nel sito FTP Verona
            If Not downloadSftpDhlVerona() Then
                wl("Errore. Download file elaborati da DHL terminato con errori")
            End If
            'Pausa 1 minuto 
            Sleep(60000)

            If Not MotoControlloDocumentiCaricoMagDhl() Then
                wl("Errore. Elaborazione file rientro controllo carico moto magazzino DHL.")
            End If

            If Not MonopattiniControlloDocumentiCaricoMagDhl() Then
                wl("Errore. Elaborazione file rientro controllo carico monopattini magazzino DHL.")
            End If

            If Not MotoControlloDocumentiSpedizioniDhl() Then
                wl("Errore. Elaborazione file rientro controllo spedizioni moto da magazzino DHL.")
            End If

            If Not MonopattiniControlloDocumentiSpedizioniDhl() Then
                wl("Errore. Elaborazione file rientro controllo spedizioni monopattini da magazzino DHL.")
            End If

            If Not MercatiControlloDocumentiB2bB2c() Then
                wl("Errore. Elaborazione file rientro controllo trasferimento monopattini da B2B a B2C.")
            End If

            wl("Elaborazione terminata.")
            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(GestioneFanticDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

#Region "CARICO MAG. DHL VERONA"

    Private Function MotoEsportaCsvCaricoMagDhlVerona() As Boolean
        Try

            Dim nomeCompletoFileLocale As String = String.Empty
            Dim nomeCompletoFileRemoto As String = String.Empty
            Dim nomeCompletoPathRemoto As String = String.Empty
            Dim nomeFile As String = String.Empty

            Dim strOut As List(Of String) = Nothing
            Dim sStr As StringBuilder = Nothing


            Dim listaBolleDhlVerona As List(Of String)
            Dim listaCarichiMagDhlVr As List(Of RDM_DHL_MotoDocCaricoMagDhlVr)

            Dim listaCarichiDaElaborare As List(Of CaricoMagDhl) = Nothing
            Dim objTempCarico As CaricoMagDhl = Nothing

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaBolleDhlVerona = dbBusiness.RDM_DHL_MotoDocCaricoMagDhlVr.Select(Function(x) x.tm_numpar).Distinct().ToList()

            End Using

            If listaBolleDhlVerona.Count < 1 Then
                Return True
            End If

            For Each BollaCamion As String In listaBolleDhlVerona

                nomeFile = "MOTOCamion" + BollaCamion + "_" + Date.Now.ToString("yyyyMMddHHmmss") & ".csv"
                nomeCompletoFileLocale = Path.Combine(_pathLocaleCSV, nomeFile)
                nomeCompletoPathRemoto = Path.Combine(_pathInputRemotoCSV, PATH_USCITA_MOTO_DA_FANTIC)
                nomeCompletoFileRemoto = Path.Combine(nomeCompletoPathRemoto, nomeFile)

                listaCarichiDaElaborare = New List(Of CaricoMagDhl)()

                strOut = New List(Of String)()

                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                    listaCarichiMagDhlVr = dbBusiness.RDM_DHL_MotoDocCaricoMagDhlVr.Where(Function(x) x.tm_numpar = BollaCamion).ToList()

                End Using

                If listaCarichiMagDhlVr.Count < 1 Then
                    Continue For
                End If

                For Each dr As RDM_DHL_MotoDocCaricoMagDhlVr In listaCarichiMagDhlVr
                    objTempCarico = New CaricoMagDhl()
                    objTempCarico.TipoDoc = dr.mma_tipork
                    objTempCarico.AnnoDoc = dr.mma_anno
                    objTempCarico.SerieDoc = dr.mma_serie
                    objTempCarico.NumDoc = dr.mma_numdoc
                    objTempCarico.RigaDoc = dr.mma_riga
                    objTempCarico.RigaMatr = dr.mma_rigaa
                    objTempCarico.RifBollaMin = dr.tm_numpar
                    objTempCarico.CodArt = dr.mm_codart
                    objTempCarico.Matricola = dr.mma_matric
                    listaCarichiDaElaborare.Add(objTempCarico)

                Next

                scriviFileCsvCaricoDHL(nomeCompletoFileLocale, nomeCompletoFileRemoto, listaCarichiDaElaborare)

            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MotoEsportaCsvCaricoMagDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try

    End Function

    Private Function MonopattiniEsportaCsvCaricoMagDhlVerona() As Boolean
        Try

            Dim nomeCompletoFileLocale As String = String.Empty
            Dim nomeCompletoFileRemoto As String = String.Empty
            Dim nomeCompletoPathRemoto As String = String.Empty
            Dim nomeFile As String = String.Empty

            Dim listaBolleDhlVerona As List(Of String)
            Dim listaCarichiMagDhlVr As List(Of RDM_DHL_MonopattiniDocCaricoMagDhlVr)

            Dim listaCarichiDaElaborare As List(Of CaricoMagDhl) = Nothing
            Dim objTempCarico As CaricoMagDhl = Nothing

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaBolleDhlVerona = dbBusiness.RDM_DHL_MonopattiniDocCaricoMagDhlVr.Select(Function(x) x.tm_numpar).Distinct().ToList()

            End Using

            If listaBolleDhlVerona.Count < 1 Then
                Return True
            End If

            For Each BollaCamion As String In listaBolleDhlVerona

                nomeFile = "MONOCamion" + BollaCamion + "_" + Date.Now.ToString("yyyyMMddHHmmss") & ".csv"
                nomeCompletoFileLocale = Path.Combine(_pathLocaleCSV, nomeFile)
                nomeCompletoPathRemoto = Path.Combine(_pathInputRemotoCSV, PATH_USCITA_MONO_DA_FANTIC)
                nomeCompletoFileRemoto = Path.Combine(nomeCompletoPathRemoto, nomeFile)
                listaCarichiDaElaborare = New List(Of CaricoMagDhl)()

                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                    listaCarichiMagDhlVr = dbBusiness.RDM_DHL_MonopattiniDocCaricoMagDhlVr.Where(Function(x) x.tm_numpar = BollaCamion).ToList()

                End Using

                If listaCarichiMagDhlVr.Count < 1 Then
                    Continue For
                End If

                For Each dr As RDM_DHL_MonopattiniDocCaricoMagDhlVr In listaCarichiMagDhlVr
                    objTempCarico = New CaricoMagDhl()
                    objTempCarico.TipoDoc = dr.mma_tipork
                    objTempCarico.AnnoDoc = dr.mma_anno
                    objTempCarico.SerieDoc = dr.mma_serie
                    objTempCarico.NumDoc = dr.mma_numdoc
                    objTempCarico.RigaDoc = dr.mma_riga
                    objTempCarico.RigaMatr = dr.mma_rigaa
                    objTempCarico.RifBollaMin = dr.tm_numpar
                    objTempCarico.CodArt = dr.mm_codart
                    objTempCarico.Matricola = dr.mma_matric
                    listaCarichiDaElaborare.Add(objTempCarico)

                Next


                scriviFileCsvCaricoDHL(nomeCompletoFileLocale, nomeCompletoFileRemoto, listaCarichiDaElaborare)

            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MonopattiniEsportaCsvCaricoMagDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try

    End Function

    Private Function scriviFileCsvCaricoDHL(ByVal nomeCompletoFileLocale As String, ByVal nomeCompletoFileRemoto As String,
                                            listaCarichiDhl As List(Of CaricoMagDhl)) As Boolean
        Try

            Dim keyRef As String = String.Empty
            Dim strOut As List(Of String) = Nothing
            Dim sStr As StringBuilder = Nothing

            strOut = New List(Of String)()
            strOut.Add("RifKey|RifBolla|CodArt|Matricola")

            For Each rigaCarico As CaricoMagDhl In listaCarichiDhl

                sStr = New StringBuilder()
                keyRef = NTSCStr(rigaCarico.TipoDoc.Trim()) + NTSCStr(rigaCarico.AnnoDoc) + NTSCStr(rigaCarico.SerieDoc.Trim()) + NTSCStr(rigaCarico.NumDoc.ToString("D6")) + NTSCStr(rigaCarico.RigaDoc.ToString("D3")) + NTSCStr(rigaCarico.RigaMatr.ToString("D3"))
                sStr.Append(keyRef)
                sStr.Append("|")
                sStr.Append(rigaCarico.RifBollaMin)
                sStr.Append("|")
                sStr.Append(rigaCarico.CodArt.Trim())
                sStr.Append("|")
                sStr.Append(rigaCarico.Matricola)
                strOut.Add(sStr.ToString())

            Next

            If strOut.Count < 1 Then
                Return True
            End If

            Using writer As StreamWriter = New StreamWriter(nomeCompletoFileLocale, True, Encoding.ASCII)
                For Each linea As String In strOut
                    writer.WriteLine(linea)
                Next
            End Using

            If Not File.Exists(nomeCompletoFileLocale) Then
                wl($"File {nomeCompletoFileLocale} NON TROVATO")
                Return False
            End If

            Dim CaricoMono As New CaricoMagDhl
            If uploadSftpDhlVerona(nomeCompletoFileLocale, nomeCompletoFileRemoto) Then
                For Each caricoInviatoADhl As CaricoMagDhl In listaCarichiDhl
                    InsertIntoHHINMAGDHL(caricoInviatoADhl)
                Next
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(scriviFileCsvCaricoDHL)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function InsertIntoHHINMAGDHL(ByVal dr As CaricoMagDhl) As Boolean
        Try

            Dim nuovoObj As New HHINMAGDHL

            Using _dbBusiness As New BusinessDbDataContext

                nuovoObj.codditt = _dittaCorrente
                nuovoObj.in_tipork = dr.TipoDoc
                nuovoObj.in_anno = dr.AnnoDoc
                nuovoObj.in_serie = dr.SerieDoc
                nuovoObj.in_numdoc = dr.NumDoc
                nuovoObj.in_riga = dr.RigaDoc
                nuovoObj.in_rigaa = dr.RigaMatr
                _dbBusiness.HHINMAGDHL.InsertOnSubmit(nuovoObj)
                _dbBusiness.SubmitChanges()

            End Using

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(InsertIntoHHINMAGDHL)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Overridable Function MotoControlloDocumentiCaricoMagDhl() As Boolean
        Try

            Dim arrayFile As String()
            Dim listaFileDaElaborare As New List(Of String)
            Dim pathLocale As String = My.Settings.PathCSV

            arrayFile = Directory.GetFiles(pathLocale)
            For Each myFile As String In arrayFile

                If myFile.Contains(MOTO_CARICO) Then
                    listaFileDaElaborare.Add(myFile)
                End If

            Next

            For Each fileDaElaborare As String In listaFileDaElaborare
                If ElaboraFileRientroCarichiMagDhl(fileDaElaborare) Then
                    spostaFileElaborato(fileDaElaborare)
                End If
            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MotoControlloDocumentiCaricoMagDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Overridable Function MonopattiniControlloDocumentiCaricoMagDhl() As Boolean
        Try

            Dim arrayFile As String()
            Dim listaFileDaElaborare As New List(Of String)
            Dim pathLocale As String = My.Settings.PathCSV

            arrayFile = Directory.GetFiles(pathLocale)
            For Each myFile As String In arrayFile

                If myFile.Contains(MONO_CARICO) Then
                    listaFileDaElaborare.Add(myFile)
                End If

            Next

            For Each fileDaElaborare As String In listaFileDaElaborare
                If ElaboraFileRientroCarichiMagDhl(fileDaElaborare) Then
                    spostaFileElaborato(fileDaElaborare)
                End If
            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MonopattiniControlloDocumentiCaricoMagDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function ElaboraFileRientroCarichiMagDhl(ByVal fileRientroCaricoMagVerona As String) As Boolean
        Try

            Dim listaCaricoVerona As New List(Of CaricoMagDhl)
            Dim objTmpCarico As CaricoMagDhl = Nothing
            Dim rigaLetta As String = String.Empty

            Dim spezzaRiga As String()
            Dim codArtTmp As String
            Dim operazioneOk As Boolean = True

            'leggo il file e lo metto un una lista
            Using reader As StreamReader = New StreamReader(fileRientroCaricoMagVerona)
                While Not reader.EndOfStream
                    rigaLetta = reader.ReadLine

                    If rigaLetta.Trim() = String.Empty Then
                        Continue While
                    End If

                    If rigaLetta.Contains(INTESTAZIONE_CAMION_CARICO) Then
                        Continue While
                    End If

                    spezzaRiga = rigaLetta.Split("|"c)
                    objTmpCarico = New CaricoMagDhl()

                    If spezzaRiga(0).Length = 18 Then
                        objTmpCarico.DocCompleto = spezzaRiga(0).Substring(0, 12)
                        objTmpCarico.TipoDoc = spezzaRiga(0).Substring(0, 1)
                        objTmpCarico.AnnoDoc = spezzaRiga(0).Substring(1, 4)
                        objTmpCarico.SerieDoc = spezzaRiga(0).Substring(5, 1)
                        objTmpCarico.NumDoc = spezzaRiga(0).Substring(6, 6)
                        objTmpCarico.RigaDoc = spezzaRiga(0).Substring(12, 3)
                        objTmpCarico.RigaMatr = spezzaRiga(0).Substring(15, 3)
                    ElseIf spezzaRiga(0).Length = 17 Then
                        objTmpCarico.DocCompleto = spezzaRiga(0).Substring(0, 12)
                        objTmpCarico.TipoDoc = spezzaRiga(0).Substring(0, 1)
                        objTmpCarico.AnnoDoc = spezzaRiga(0).Substring(1, 4)
                        objTmpCarico.SerieDoc = " "
                        objTmpCarico.NumDoc = spezzaRiga(0).Substring(5, 6)
                        objTmpCarico.RigaDoc = spezzaRiga(0).Substring(11, 3)
                        objTmpCarico.RigaMatr = spezzaRiga(0).Substring(14, 3)
                    Else
                        wl("Formato chiave non valido")
                        operazioneOk = False
                        Continue While
                    End If


                    objTmpCarico.RifBollaMin = spezzaRiga(1)
                    codArtTmp = spezzaRiga(2)
                    If Not oMenu.ValCodiceDb(codArtTmp, _dittaCorrente, "ARTICO", "S") Then
                        wl($"Articolo: {codArtTmp} di ritorno non trovato")
                        operazioneOk = False
                    End If

                    objTmpCarico.CodArt = codArtTmp
                    objTmpCarico.Matricola = spezzaRiga(3)
                    listaCaricoVerona.Add(objTmpCarico)

                End While
            End Using

            If Not operazioneOk Then
                wl($"File: {fileRientroCaricoMagVerona} non elaborato.")
                Return False
            End If

            'creo una lista con solo i documenti da scorrere
            Dim listDoc As List(Of String)
            listDoc = listaCaricoVerona.Select(Function(x) x.DocCompleto).Distinct.ToList

            Dim listaDoc As List(Of CaricoMagDhl) = Nothing
            For Each doc As String In listDoc

                listaDoc = listaCaricoVerona.Where(Function(x) x.DocCompleto = doc).ToList()
                If creaDocNeutroDaRilevazioneDHL(listaDoc, doc) Then
                    operazioneOk = False
                End If

            Next

            Return operazioneOk

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(ElaboraFileRientroCarichiMagDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "USCITE MAG. DHL VERONA"

    Private Function MotoEsportCsvUsciteDaDhlVerona() As Boolean
        Try

            Dim bErrori As Boolean = False
            Dim nomeCompletoFileRemoto As String = String.Empty
            Dim nomeCompletoFile As String = String.Empty
            Dim nomeCompletoPathRemoto As String = String.Empty
            Dim strOut As List(Of String) = Nothing
            Dim sStr As StringBuilder = Nothing
            Dim dttAngra As New DataTable
            Dim dttDestDiv As New DataTable
            Dim dttStato As New DataTable
            Dim listaDocUscitaDhlVr As List(Of RDM_DHL_MotoDocUscitaMagDhlVr)

            Dim nomeFile As String = "fmspedMB_" + Date.Now.ToString("yyyyMMddHHmmss") & ".csv"
            nomeCompletoFile = Path.Combine(_pathLocaleCSV, nomeFile)
            nomeCompletoPathRemoto = Path.Combine(_pathInputRemotoCSV, PATH_USCITA_MOTO_DA_FANTIC)
            nomeCompletoFileRemoto = Path.Combine(nomeCompletoPathRemoto, nomeFile)
            strOut = New List(Of String)()

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaDocUscitaDhlVr = dbBusiness.RDM_DHL_MotoDocUscitaMagDhlVr.ToList()

            End Using

            Dim keyRef As String = String.Empty
            Dim codStatoSped As String = String.Empty

            Dim codSpedizioniItalia As String = My.Settings.codSpedizioniItalia
            Dim codSpedizioniEsteroMoto As String = My.Settings.codSpedizioniEsteroMoto
            Dim codSpedizioniEsteroMonopattini As String = My.Settings.codSpedizioniEsteroMonopattini
            Dim codSpedizioniEsteroBici As String = My.Settings.codSpedizioniEsteroBici


            strOut.Add("RifKey|RagSociale|Indirizzo|Cap|Localita|StatoIso|NumColli|Peso|Volume|Note|IdMail|DataCons|Telefono|SpondaIdraulica|CodAccuntDhl|CodArt|BB")

            Dim usciteMoto As New SpedizioniMagDhl
            For Each dr As RDM_DHL_MotoDocUscitaMagDhlVr In listaDocUscitaDhlVr

                codStatoSped = String.Empty
                sStr = New StringBuilder()
                sStr.Append(NTSCStr(dr.mm_tipork).Trim() + NTSCStr(dr.mm_anno) + NTSCStr(dr.mm_serie).Trim() + NTSCStr(dr.mm_numdoc).PadLeft(6, "0"c) + NTSCStr(dr.mm_riga).PadLeft(3, "0"c))
                sStr.Append("|")

                usciteMoto.Coddest = dr.tm_coddest
                usciteMoto.Conto = dr.tm_conto

                If NTSCInt(dr.tm_coddest) > 0 Then

                    If Not scriviFileSpedizioneDestDiv(usciteMoto, dttDestDiv, codStatoSped, sStr) Then
                        bErrori = True
                        Continue For
                    End If

                Else

                    If Not scriviFileSpedSedePrincipale(usciteMoto, dttAngra, codStatoSped, sStr) Then
                        bErrori = True
                        Continue For
                    End If

                End If

                If Not oMenu.ValCodiceDb(codStatoSped.Trim(), _dittaCorrente, "TABSTAT", "S", "", dttStato) Then
                    wl(String.Format("ATTENZIONE codice nazione {0} non trovato", codStatoSped.Trim()))
                    bErrori = True
                    Continue For
                End If

                sStr.Append("|")
                Dim statoTmp As String = ""
                If Not dttStato Is Nothing Then
                    If dttStato.Rows.Count = 1 Then
                        statoTmp = NTSCStr(dttStato.Rows(0)!tb_siglaiso)
                    End If
                End If
                sStr.Append(statoTmp)

                sStr.Append("|")
                sStr.Append(1)
                sStr.Append("|") 'peso
                sStr.Append(NTSCStr(dr.ar_pesolor))
                sStr.Append("|0") 'volume
                sStr.Append("|") 'note
                sStr.Append(NTSCStr(dr.tm_hhnotedhl))
                sStr.Append("|")

                'Seleziono solo la prima e mail 
                Dim divEmail As String()

                If NTSCInt(dr.tm_coddest) > 0 Then
                    divEmail = NTSCStr(dttDestDiv.Rows(0)!dd_email).Split(";"c)
                Else
                    divEmail = NTSCStr(dttAngra.Rows(0)!an_email).Split(";"c)
                End If


                Dim tmpEmail As String = String.Empty
                If divEmail.Length > 0 Then
                    tmpEmail = divEmail(0)
                End If
                If tmpEmail.Length > 70 Then
                    sStr.Append(tmpEmail.Substring(0, 70))
                Else
                    sStr.Append(tmpEmail)
                End If
                sStr.Append("|")

                Dim organig As New organig
                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)
                    organig = dbBusiness.organig.Where(Function(x) x.og_conto = dr.tm_conto And x.og_codruaz = "BO").FirstOrDefault()
                End Using

                Dim numTelef As String = String.Empty
                If Not organig Is Nothing Then
                    If Not organig.og_telef Is Nothing Then
                        numTelef = organig.og_telef
                    End If
                End If
                sStr.Append(numTelef)
                sStr.Append("|")

                sStr.Append("|SI")
                sStr.Append("|")

                If codStatoSped = "I" Then
                    sStr.Append(codSpedizioniItalia)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00010" Then
                    sStr.Append(codSpedizioniEsteroMoto)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00020" Then
                    sStr.Append(codSpedizioniEsteroBici)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00030" Then
                    sStr.Append(codSpedizioniEsteroMonopattini)
                Else
                    sStr.Append("00000000000")
                End If

                sStr.Append("|")
                sStr.Append(NTSCStr(dr.mm_codart))
                sStr.Append("|B2B")
                strOut.Add(sStr.ToString())

                'Se riga con QTA maggiore a 1 replico la riga per la QTA
                If dr.mm_quant > 1 Then
                    For cont As Integer = 1 To dr.mm_quant - 1
                        strOut.Add(sStr.ToString())
                    Next
                End If
            Next

            If strOut.Count < 2 Then
                Return True
            End If

            Using writer As StreamWriter = New StreamWriter(nomeCompletoFile, True, Encoding.ASCII)
                For Each linea As String In strOut
                    writer.WriteLine(linea)
                Next
            End Using

            If Not File.Exists(nomeCompletoFile) Then
                wl($"File {nomeCompletoFile} NON TROVATO")
                Return False
            End If

            If uploadSftpDhlVerona(nomeCompletoFile, nomeCompletoFileRemoto) Then
                Dim motoToInsert As New SpedizioniMagDhl
                For Each dr As RDM_DHL_MotoDocUscitaMagDhlVr In listaDocUscitaDhlVr
                    motoToInsert.TipoDoc = dr.mm_tipork
                    motoToInsert.AnnoDoc = dr.mm_anno
                    motoToInsert.SerieDoc = dr.mm_serie
                    motoToInsert.NumDoc = dr.mm_numdoc
                    motoToInsert.RigaDoc = dr.mm_riga
                    InsertIntoHHOUTMAGDHL(motoToInsert)
                Next
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MotoEsportCsvUsciteDaDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Private Function MonopattiniEsportCsvUsciteDaDhlVerona() As Boolean
        Try
            Dim bErrori As Boolean = False
            Dim nomeCompletoFileRemoto As String = String.Empty
            Dim nomeCompletoFile As String = String.Empty
            Dim nomeCompletoPathRemoto As String = String.Empty
            Dim strOut As List(Of String) = Nothing
            Dim sStr As StringBuilder = Nothing
            Dim dttAngra As New DataTable
            Dim dttDestDiv As New DataTable
            Dim dttStato As New DataTable
            Dim listaDocUscitaDhlVr As List(Of RDM_DHL_MonopattiniDocUscitaMagB2BDhlVr)

            Dim nomeFile As String = "fmspedB2B_" + Date.Now.ToString("yyyyMMddHHmmss") & ".csv"
            nomeCompletoFile = Path.Combine(_pathLocaleCSV, nomeFile)
            nomeCompletoPathRemoto = Path.Combine(_pathInputRemotoCSV, PATH_USCITA_MONO_DA_FANTIC)
            nomeCompletoFileRemoto = Path.Combine(nomeCompletoPathRemoto, nomeFile)
            strOut = New List(Of String)()

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaDocUscitaDhlVr = dbBusiness.RDM_DHL_MonopattiniDocUscitaMagB2BDhlVr.ToList()

            End Using

            Dim keyRef As String = String.Empty
            Dim codStatoSped As String = String.Empty

            Dim codSpedizioniItalia As String = My.Settings.codSpedizioniItalia
            Dim codSpedizioniEsteroMoto As String = My.Settings.codSpedizioniEsteroMoto
            Dim codSpedizioniEsteroMonopattini As String = My.Settings.codSpedizioniEsteroMonopattini
            Dim codSpedizioniEsteroBici As String = My.Settings.codSpedizioniEsteroBici


            strOut.Add("RifKey|RagSociale|Indirizzo|Cap|Localita|StatoIso|NumColli|Peso|Volume|Note|IdMail|DataCons|Telefono|SpondaIdraulica|CodAccuntDhl|CodArt|BB")

            Dim usciteMono As New SpedizioniMagDhl

            For Each dr As RDM_DHL_MonopattiniDocUscitaMagB2BDhlVr In listaDocUscitaDhlVr

                codStatoSped = String.Empty
                sStr = New StringBuilder()
                sStr.Append(NTSCStr(dr.mm_tipork).Trim() + NTSCStr(dr.mm_anno) + NTSCStr(dr.mm_serie).Trim() + NTSCStr(dr.mm_numdoc).PadLeft(6, "0"c) + NTSCStr(dr.mm_riga).PadLeft(3, "0"c))
                sStr.Append("|")

                usciteMono.Coddest = dr.tm_coddest
                usciteMono.Conto = dr.tm_conto

                If NTSCInt(dr.tm_coddest) > 0 Then

                    If Not scriviFileSpedizioneDestDiv(usciteMono, dttDestDiv, codStatoSped, sStr) Then
                        bErrori = True
                        Continue For
                    End If

                Else

                    If Not scriviFileSpedSedePrincipale(usciteMono, dttAngra, codStatoSped, sStr) Then
                        bErrori = True
                        Continue For
                    End If

                End If

                If Not oMenu.ValCodiceDb(codStatoSped.Trim(), _dittaCorrente, "TABSTAT", "S", "", dttStato) Then
                    wl(String.Format("ATTENZIONE codice nazione {0} non trovato", codStatoSped.Trim()))
                    bErrori = True
                    Continue For
                End If

                sStr.Append("|")
                Dim statoTmp As String = ""
                If Not dttStato Is Nothing Then
                    If dttStato.Rows.Count = 1 Then
                        statoTmp = NTSCStr(dttStato.Rows(0)!tb_siglaiso)
                    End If
                End If
                sStr.Append(statoTmp)

                sStr.Append("|")
                sStr.Append(1)
                sStr.Append("|") 'peso
                sStr.Append(NTSCStr(dr.ar_pesolor))
                sStr.Append("|0") 'volume
                sStr.Append("|") 'note
                sStr.Append(NTSCStr(dr.tm_hhnotedhl))
                sStr.Append("|")

                'Seleziono solo la prima e mail 
                Dim divEmail As String()

                If NTSCInt(dr.tm_coddest) > 0 Then
                    divEmail = NTSCStr(dttDestDiv.Rows(0)!dd_email).Split(";"c)
                Else
                    divEmail = NTSCStr(dttAngra.Rows(0)!an_email).Split(";"c)
                End If


                Dim tmpEmail As String = String.Empty
                If divEmail.Length > 0 Then
                    tmpEmail = divEmail(0)
                End If
                If tmpEmail.Length > 70 Then
                    sStr.Append(tmpEmail.Substring(0, 70))
                Else
                    sStr.Append(tmpEmail)
                End If
                sStr.Append("|")

                Dim organig As New organig
                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)
                    organig = dbBusiness.organig.Where(Function(x) x.og_conto = dr.tm_conto And x.og_codruaz = "BO").FirstOrDefault()
                End Using

                Dim numTelef As String = String.Empty
                If Not organig Is Nothing Then
                    If Not organig.og_telef Is Nothing Then
                        numTelef = organig.og_telef
                    End If
                End If
                sStr.Append(numTelef)
                sStr.Append("|")

                sStr.Append("|SI")
                sStr.Append("|")

                If codStatoSped = "I" Then
                    sStr.Append(codSpedizioniItalia)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00010" Then
                    sStr.Append(codSpedizioniEsteroMoto)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00020" Then
                    sStr.Append(codSpedizioniEsteroBici)
                ElseIf codStatoSped <> "I" And NTSCStr(dr.ar_codcla1) = "00030" Then
                    sStr.Append(codSpedizioniEsteroMonopattini)
                Else
                    sStr.Append("00000000000")
                End If

                sStr.Append("|")
                sStr.Append(NTSCStr(dr.mm_codart))

                sStr.Append("|B2C")

                strOut.Add(sStr.ToString())

                'Se riga con QTA maggiore a 1 replico la riga per la QTA
                If dr.mm_quant > 1 Then
                    For cont As Integer = 1 To dr.mm_quant - 1
                        strOut.Add(sStr.ToString())
                    Next
                End If
            Next

            If strOut.Count < 2 Then
                Return True
            End If

            Using writer As StreamWriter = New StreamWriter(nomeCompletoFile, True, Encoding.ASCII)
                For Each linea As String In strOut
                    writer.WriteLine(linea)
                Next
            End Using

            If Not File.Exists(nomeCompletoFile) Then
                wl($"File {nomeCompletoFile} NON TROVATO")
                Return False
            End If

            If uploadSftpDhlVerona(nomeCompletoFile, nomeCompletoFileRemoto) Then
                Dim monopattinoToInsert As New SpedizioniMagDhl
                For Each dr As RDM_DHL_MonopattiniDocUscitaMagB2BDhlVr In listaDocUscitaDhlVr
                    monopattinoToInsert.TipoDoc = dr.mm_tipork
                    monopattinoToInsert.AnnoDoc = dr.mm_anno
                    monopattinoToInsert.SerieDoc = dr.mm_serie
                    monopattinoToInsert.NumDoc = dr.mm_numdoc
                    monopattinoToInsert.RigaDoc = dr.mm_riga
                    InsertIntoHHOUTMAGDHL(monopattinoToInsert)
                Next
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MonopattiniEsportCsvUsciteDaDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Public Function scriviFileSpedizioneDestDiv(ByVal spedizDhl As SpedizioniMagDhl, ByRef dttDestDiv As DataTable,
                                                  ByRef codStatoSped As String, ByRef sStr As StringBuilder) As Boolean
        Try

            If Not oMenu.ValCodiceDb(NTSCStr(spedizDhl.Coddest).Trim(), _dittaCorrente, "DESTDIV", "N", "", dttDestDiv, NTSCStr(spedizDhl.Conto).Trim()) Then
                wl(String.Format("ATTENZIONE destinazione diversa {0} del cliente cod. {0} non trovato", NTSCStr(spedizDhl.Coddest), NTSCStr(spedizDhl.Conto).Trim()))
                Return False
            End If

            If Not dttDestDiv Is Nothing Then
                If dttDestDiv.Rows.Count = 1 Then
                    If Not dttDestDiv.Rows(0)!dd_stato Is Nothing Then
                        If NTSCStr(dttDestDiv.Rows(0)!dd_stato).Trim() <> String.Empty Then
                            codStatoSped = NTSCStr(dttDestDiv.Rows(0)!dd_stato).Trim()
                        End If
                    End If
                End If
            End If

            If codStatoSped.Trim() = String.Empty Then
                If Not dttDestDiv Is Nothing Then
                    If dttDestDiv.Rows.Count = 1 Then
                        If Not dttDestDiv.Rows(0)!dd_codcomu Is Nothing Then
                            If NTSCStr(dttDestDiv.Rows(0)!dd_codcomu).Trim() <> String.Empty Then
                                If oMenu.ValCodiceDb(NTSCStr(dttDestDiv.Rows(0)!dd_codcomu).Trim(), _dittaCorrente, "COMUNI", "S") Then
                                    codStatoSped = "I"
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If NTSCStr(dttDestDiv.Rows(0)!dd_nomdest).Length > 35 Then
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_nomdest).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_nomdest))
            End If
            sStr.Append("|")
            If NTSCStr(dttDestDiv.Rows(0)!dd_inddest).Length > 35 Then
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_inddest).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_inddest))
            End If
            sStr.Append("|")
            sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_capdest))
            sStr.Append("|")
            If NTSCStr(dttDestDiv.Rows(0)!dd_locdest).Length > 35 Then
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_locdest).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttDestDiv.Rows(0)!dd_locdest))
            End If
            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(scriviFileSpedizioneDestDiv)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Public Function scriviFileSpedSedePrincipale(ByVal dr As SpedizioniMagDhl, ByRef dttAngra As DataTable,
                                                       ByRef codStatoSped As String, ByRef sStr As StringBuilder) As Boolean
        Try
            If Not oMenu.ValCodiceDb(NTSCStr(dr.Conto).Trim(), _dittaCorrente, "ANAGRA", "S", "", dttAngra) Then
                wl(String.Format("ATTENZIONE cliente {0} non trovato", NTSCStr(dr.Conto).Trim()))
                Return False
            End If

            If Not dttAngra Is Nothing Then
                If dttAngra.Rows.Count = 1 Then
                    If Not dttAngra.Rows(0)!an_stato Is Nothing Then
                        If NTSCStr(dttAngra.Rows(0)!an_stato).Trim() <> String.Empty Then
                            codStatoSped = NTSCStr(dttAngra.Rows(0)!an_stato).Trim()
                        End If
                    End If
                End If
            End If


            If codStatoSped.Trim() = String.Empty Then
                If Not dttAngra Is Nothing Then
                    If dttAngra.Rows.Count = 1 Then
                        If Not dttAngra.Rows(0)!an_codcomu Is Nothing Then
                            If NTSCStr(dttAngra.Rows(0)!an_codcomu).Trim() <> String.Empty Then
                                If oMenu.ValCodiceDb(NTSCStr(dttAngra.Rows(0)!an_codcomu).Trim(), _dittaCorrente, "COMUNI", "S") Then
                                    codStatoSped = "I"
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If NTSCStr(dttAngra.Rows(0)!an_descr1).Length > 35 Then
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_descr1).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_descr1))
            End If
            sStr.Append("|")
            If NTSCStr(dttAngra.Rows(0)!an_indir).Length > 35 Then
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_indir).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_indir))
            End If
            sStr.Append("|")
            sStr.Append(NTSCStr(dttAngra.Rows(0)!an_cap))
            sStr.Append("|")
            If NTSCStr(dttAngra.Rows(0)!an_citta).Length > 35 Then
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_citta).Substring(0, 35))
            Else
                sStr.Append(NTSCStr(dttAngra.Rows(0)!an_citta))
            End If
            Return True
        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(scriviFileSpedSedePrincipale)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Public Function InsertIntoHHOUTMAGDHL(ByVal dr As SpedizioniMagDhl) As Boolean
        Try

            Dim nuovoObj As New HHOUTMAGDHL

            Using _dbBusiness As New BusinessDbDataContext

                nuovoObj.codditt = _dittaCorrente
                nuovoObj.ou_tipork = dr.TipoDoc
                nuovoObj.ou_anno = dr.AnnoDoc
                nuovoObj.ou_serie = dr.SerieDoc
                nuovoObj.ou_numdoc = dr.NumDoc
                nuovoObj.ou_riga = dr.RigaDoc
                _dbBusiness.HHOUTMAGDHL.InsertOnSubmit(nuovoObj)
                _dbBusiness.SubmitChanges()

            End Using

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(InsertIntoHHOUTMAGDHL)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Public Function MotoControlloDocumentiSpedizioniDhl() As Boolean
        Try

            Dim arrayFile As String()
            Dim listaFileDaElaborare As New List(Of String)
            Dim pathLocale As String = My.Settings.PathCSV

            arrayFile = Directory.GetFiles(pathLocale)
            For Each myFile As String In arrayFile

                If myFile.Contains(MOTO_SPED) Then
                    listaFileDaElaborare.Add(myFile)
                End If

            Next

            For Each fileDaElaborare As String In listaFileDaElaborare

                If ElaboraFileRientroUsciteMagDhl(fileDaElaborare) Then
                    spostaFileElaborato(fileDaElaborare)
                End If

            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MotoControlloDocumentiSpedizioniDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Public Function MonopattiniControlloDocumentiSpedizioniDhl() As Boolean
        Try

            Dim arrayFile As String()
            Dim listaFileDaElaborare As New List(Of String)
            Dim pathLocale As String = My.Settings.PathCSV

            arrayFile = Directory.GetFiles(pathLocale)
            For Each myFile As String In arrayFile

                If myFile.Contains(MONO_SPED) Then
                    listaFileDaElaborare.Add(myFile)
                End If

            Next

            For Each fileDaElaborare As String In listaFileDaElaborare

                If ElaboraFileRientroUsciteMagDhl(fileDaElaborare) Then
                    spostaFileElaborato(fileDaElaborare)
                End If

            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MonopattiniControlloDocumentiSpedizioniDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function
    Private Function ElaboraFileRientroUsciteMagDhl(ByVal fileRientroSpedizioneMagVerona As String) As Boolean
        Try

            Dim listaSpedizioneVerona As New List(Of SpedizioniMagDhl)
            Dim objTmpSpedizione As SpedizioniMagDhl = Nothing
            Dim rigaLetta As String = String.Empty

            Dim spezzaRiga As String()
            Dim codArtTmp As String
            Dim operazioneOk As Boolean = True

            'leggo il file e lo metto un una lista
            Using reader As StreamReader = New StreamReader(fileRientroSpedizioneMagVerona)
                While Not reader.EndOfStream
                    rigaLetta = reader.ReadLine

                    If rigaLetta.Trim() = String.Empty Then
                        Continue While
                    End If

                    If rigaLetta.Contains(INTESTAZIONE_SPEDIZIONE) Then
                        Continue While
                    End If

                    spezzaRiga = rigaLetta.Split("|"c)
                    objTmpSpedizione = New SpedizioniMagDhl()

                    If spezzaRiga(0).Length = 15 Then
                        objTmpSpedizione.DocCompleto = spezzaRiga(0).Substring(0, 12)
                        objTmpSpedizione.TipoDoc = spezzaRiga(0).Substring(0, 1)
                        objTmpSpedizione.AnnoDoc = spezzaRiga(0).Substring(1, 4)
                        objTmpSpedizione.SerieDoc = spezzaRiga(0).Substring(5, 1)
                        objTmpSpedizione.NumDoc = spezzaRiga(0).Substring(6, 6)
                        objTmpSpedizione.RigaDoc = spezzaRiga(0).Substring(12, 3)
                    ElseIf spezzaRiga(0).Length = 14 Then
                        objTmpSpedizione.DocCompleto = spezzaRiga(0).Substring(0, 11)
                        objTmpSpedizione.TipoDoc = spezzaRiga(0).Substring(0, 1)
                        objTmpSpedizione.AnnoDoc = spezzaRiga(0).Substring(1, 4)
                        objTmpSpedizione.SerieDoc = " "
                        objTmpSpedizione.NumDoc = spezzaRiga(0).Substring(5, 6)
                        objTmpSpedizione.RigaDoc = spezzaRiga(0).Substring(11, 3)
                    Else
                        wl("Tracciato documento riga ingresso non valido")
                        operazioneOk = False
                        Continue While
                    End If

                    codArtTmp = spezzaRiga(15)
                    If Not oMenu.ValCodiceDb(codArtTmp, _dittaCorrente, "ARTICO", "S") Then
                        wl($"Articolo: {codArtTmp} di ritorno non trovato")
                        operazioneOk = False
                        Continue While
                    End If

                    objTmpSpedizione.CodArt = codArtTmp
                    objTmpSpedizione.Matricola = spezzaRiga(17)
                    listaSpedizioneVerona.Add(objTmpSpedizione)

                End While
            End Using

            If Not operazioneOk Then
                wl($"File {fileRientroSpedizioneMagVerona} non elaborato.")
                Return False
            End If

            'creo una lista con solo i documenti da scorrere
            Dim listDoc As List(Of String)
            listDoc = listaSpedizioneVerona.Select(Function(x) x.DocCompleto).Distinct.ToList

            Dim righeDocumento As List(Of SpedizioniMagDhl) = Nothing
            Dim sMsgErr As StringBuilder
            For Each docCompleto As String In listDoc

                sMsgErr = New StringBuilder()
                righeDocumento = listaSpedizioneVerona.Where(Function(x) x.DocCompleto = docCompleto).ToList
                If Not aggiornaDocSpedizioniMagDHLConMatricole(righeDocumento, docCompleto, sMsgErr) Then
                    sendMailErroreMatricoleNotaPrelievo(sMsgErr, righeDocumento.FirstOrDefault())
                    operazioneOk = False
                End If

            Next

            Return operazioneOk

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(ElaboraFileRientroUsciteMagDhl)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "TRASF. FROM B2B TO B2C"

    Public Function MercatiDocCaricoB2bB2c() As Boolean
        Try

            Dim listaArticoliNotePrelievo As New List(Of String)
            Dim nomeCompletoPathRemoto As String = String.Empty
            Dim nomeCompletoFileRemoto As String = String.Empty
            Dim nomeCompletoFile As String = String.Empty
            Dim strOut As List(Of String) = Nothing
            Dim sStr As StringBuilder = Nothing
            Dim nomeFile As String = "fmassign_" + Date.Now.ToString("yyyyMMddHHmmss") & ".csv"

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaArticoliNotePrelievo = dbBusiness.RDM_DHL_MonopattiniDocTrasfFromB2BToB2C.Select(Function(x) x.mm_codart).Distinct.ToList

            End Using

            If listaArticoliNotePrelievo.Count < 1 Then
                Return True
            End If


            nomeCompletoFile = Path.Combine(_pathLocaleCSV, nomeFile)
            nomeCompletoPathRemoto = Path.Combine(_pathInputRemotoCSV, PATH_USCITA_MONO_DA_FANTIC)
            nomeCompletoFileRemoto = Path.Combine(nomeCompletoPathRemoto, nomeFile)

            strOut = New List(Of String)()
            strOut.Add("CodArt|Mercato|Quant")

            Dim listaArticoliStock As List(Of String) = Nothing

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaArticoliStock = dbBusiness.RDM_DHL_StockMonopattini.Select(Function(x) x.km_codart).Distinct.ToList

            End Using

            For Each codArt As String In listaArticoliStock

                Dim listaStockMono As New List(Of RDM_DHL_StockMonopattini)
                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                    listaStockMono = dbBusiness.RDM_DHL_StockMonopattini.Where(Function(x) x.km_codart = codArt).ToList

                End Using

                Dim quantB2B As Decimal = 0
                Dim quantB2C As Decimal = 0

                For Each stock As RDM_DHL_StockMonopattini In listaStockMono

                    If (stock.km_magaz = _magzzinoB2CDhl) Then
                        quantB2C += stock.Giacenza
                    Else
                        quantB2B += stock.Giacenza
                    End If

                Next

                Dim listaNoteDaElaborare As New List(Of RDM_DHL_MonopattiniDocTrasfFromB2BToB2C)

                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                    listaNoteDaElaborare = dbBusiness.RDM_DHL_MonopattiniDocTrasfFromB2BToB2C.Where(Function(x) x.mm_codart = codArt).ToList

                End Using

                For Each dr As RDM_DHL_MonopattiniDocTrasfFromB2BToB2C In listaNoteDaElaborare

                    InsertIntoHHB2CATTMAGDHL(dr)
                    quantB2C += NTSCDec(dr.mm_quant)
                    quantB2B -= NTSCDec(dr.mm_quant)

                Next

                sStr = New StringBuilder()
                sStr.Append(codArt)
                sStr.Append("|")
                sStr.Append(MERCATO_B2C)
                sStr.Append("|")
                sStr.Append(NTSCStr(quantB2C))
                strOut.Add(sStr.ToString())

                sStr = New StringBuilder()
                sStr.Append(codArt)
                sStr.Append("|")
                sStr.Append(MERCATO_B2B)
                sStr.Append("|")
                sStr.Append(NTSCStr(quantB2B))
                strOut.Add(sStr.ToString())

            Next

            Using writer As StreamWriter = New StreamWriter(nomeCompletoFile, True, Encoding.ASCII)
                For Each linea As String In strOut
                    writer.WriteLine(linea)
                Next
            End Using

            If Not File.Exists(nomeCompletoFile) Then
                wl($"File {nomeCompletoFile} NON TROVATO")
                Return False
            End If

            uploadSftpDhlVerona(nomeCompletoFile, nomeCompletoFileRemoto)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MercatiDocCaricoB2bB2c)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function InsertIntoHHB2CATTMAGDHL(ByVal dr As RDM_DHL_MonopattiniDocTrasfFromB2BToB2C) As Boolean
        Try

            Dim nuovoObj As New HHB2CATTMAGDHL

            Using _dbBusiness As New BusinessDbDataContext

                nuovoObj.codditt = _dittaCorrente
                nuovoObj.bc_tipork = dr.mm_tipork
                nuovoObj.bc_anno = dr.mm_anno
                nuovoObj.bc_serie = dr.mm_serie
                nuovoObj.bc_numdoc = dr.mm_numdoc
                nuovoObj.bc_riga = dr.mm_riga
                nuovoObj.bc_codart = dr.mm_codart
                nuovoObj.bc_quant = dr.mm_quant
                _dbBusiness.HHB2CATTMAGDHL.InsertOnSubmit(nuovoObj)
                _dbBusiness.SubmitChanges()

            End Using

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(InsertIntoHHB2CATTMAGDHL)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function MercatiControlloDocumentiB2bB2c() As Boolean
        Try

            Dim arrayFile As String()
            Dim listaFileDaElaborare As New List(Of String)
            Dim pathLocale As String = My.Settings.PathCSV

            arrayFile = Directory.GetFiles(pathLocale)
            For Each myFile As String In arrayFile

                If myFile.Contains(B2B_TRASF) Then
                    listaFileDaElaborare.Add(myFile)
                End If

            Next

            For Each fileDaElaborare As String In listaFileDaElaborare

                If ElaboraFileRientroTrasferimentiFromB2bTob2c(fileDaElaborare) Then
                    spostaFileElaborato(fileDaElaborare)
                End If

            Next

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(MercatiControlloDocumentiB2bB2c)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function ElaboraFileRientroTrasferimentiFromB2bTob2c(ByVal fileRientroSpedizioneMagVerona As String) As Boolean
        Try

            Dim rigaLetta As String = String.Empty
            Dim spezzaRiga As String()
            Dim codArtTmp As String
            Dim operazioneOk As Boolean = True
            Dim sMsgErr As New StringBuilder()

            Dim objTmpMonopattini As MatricolaB2bB2c = Nothing
            Dim objStockInterno As RDM_DHL_StockPf = Nothing
            Dim listaMatricoleDaSpostare As New List(Of RDM_DHL_StockPf)

            'leggo il file e lo metto un una lista
            Using reader As StreamReader = New StreamReader(fileRientroSpedizioneMagVerona)
                While Not reader.EndOfStream
                    rigaLetta = reader.ReadLine

                    If rigaLetta.Trim() = String.Empty Then
                        Continue While
                    End If

                    If rigaLetta.Contains(INTESTAZIONE_TRASF_B2C) Then
                        Continue While
                    End If

                    spezzaRiga = rigaLetta.Split("|"c)
                    objTmpMonopattini = New MatricolaB2bB2c()

                    codArtTmp = spezzaRiga(0)
                    If Not oMenu.ValCodiceDb(codArtTmp, _dittaCorrente, "ARTICO", "S") Then
                        wl($"Articolo: {codArtTmp} di ritorno non trovato")
                        operazioneOk = False
                        Continue While
                    End If

                    objTmpMonopattini.CodArt = codArtTmp
                    objTmpMonopattini.Matricola = spezzaRiga(1)

                    If spezzaRiga(2).Trim() = MERCATO_B2C Then
                        objTmpMonopattini.MagMercato = _magzzinoB2CDhl
                    Else
                        objTmpMonopattini.MagMercato = _magazzinoDhl
                    End If

                    Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                        objStockInterno = dbBusiness.RDM_DHL_StockPf.Where(Function(x) x.mma_matric = objTmpMonopattini.Matricola).FirstOrDefault()

                    End Using

                    If objStockInterno Is Nothing Then
                        sMsgErr.AppendLine("ATTENZIONE Matricola non associata a B2B o B2C")
                        operazioneOk = False
                        Continue While
                    End If

                    If objStockInterno.km_magaz <> objTmpMonopattini.MagMercato Then
                        listaMatricoleDaSpostare.Add(objStockInterno)
                    End If

                End While
            End Using

            Dim listaNotePrelievoDaAggiornare As List(Of RDM_DHL_MonopattiniDocRientriFromB2BToB2C) = Nothing

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                listaNotePrelievoDaAggiornare = dbBusiness.RDM_DHL_MonopattiniDocRientriFromB2BToB2C.ToList()

            End Using

            Dim matricolaSelezionata As RDM_DHL_StockPf
            Dim listaNotePrelievoSped As New List(Of SpedizioniMagDhl)
            Dim objTmpSpedizione As SpedizioniMagDhl
            For Each notaPrelievo As RDM_DHL_MonopattiniDocRientriFromB2BToB2C In listaNotePrelievoDaAggiornare
                For qta As Integer = 1 To notaPrelievo.bc_quant

                    objTmpSpedizione = New SpedizioniMagDhl()
                    objTmpSpedizione.DocCompleto = notaPrelievo.bc_tipork.Trim() + notaPrelievo.bc_anno.ToString("D4") + notaPrelievo.bc_serie.Trim() + notaPrelievo.bc_numdoc.ToString("D6")
                    objTmpSpedizione.TipoDoc = notaPrelievo.bc_tipork
                    objTmpSpedizione.AnnoDoc = notaPrelievo.bc_anno
                    objTmpSpedizione.SerieDoc = notaPrelievo.bc_serie
                    objTmpSpedizione.NumDoc = notaPrelievo.bc_numdoc
                    objTmpSpedizione.RigaDoc = notaPrelievo.bc_riga

                    Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                        matricolaSelezionata = listaMatricoleDaSpostare.Where(Function(x) x.km_codart = notaPrelievo.bc_codart).FirstOrDefault()

                    End Using

                    If matricolaSelezionata Is Nothing Then
                        Continue For
                    End If

                    objTmpSpedizione.CodArt = matricolaSelezionata.km_codart
                    objTmpSpedizione.Matricola = matricolaSelezionata.mma_matric
                    listaNotePrelievoSped.Add(objTmpSpedizione)
                    listaMatricoleDaSpostare.Remove(matricolaSelezionata)

                Next

            Next

            'creo una lista con solo i documenti da scorrere
            Dim listDoc As List(Of String)
            listDoc = listaNotePrelievoSped.Select(Function(x) x.DocCompleto).Distinct.ToList

            Dim righeDocumento As List(Of SpedizioniMagDhl) = Nothing
            Dim tmpDocBmi As SpedizioniMagDhl
            For Each docCompleto As String In listDoc

                sMsgErr = New StringBuilder()
                righeDocumento = listaNotePrelievoSped.Where(Function(x) x.DocCompleto = docCompleto).ToList
                If Not aggiornaDocSpedizioniMagDHLConMatricole(righeDocumento, docCompleto, sMsgErr) Then
                    sendMailErroreMatricoleNotaPrelievo(sMsgErr, righeDocumento.FirstOrDefault())
                    operazioneOk = False
                    Continue For
                End If


                If Not CreaBMISpostamentoMerceDaB2bB2c(righeDocumento.FirstOrDefault(), tmpDocBmi, sMsgErr) Then
                    sendMailErroreMatricoleBmiSpostamentoB2bB2c(sMsgErr, tmpDocBmi)
                    operazioneOk = False
                    Continue For
                End If

            Next

            If operazioneOk Then
                Dim monopattinoToInsert As New SpedizioniMagDhl
                For Each dr As RDM_DHL_MonopattiniDocRientriFromB2BToB2C In listaNotePrelievoDaAggiornare
                    monopattinoToInsert.TipoDoc = dr.bc_tipork
                    monopattinoToInsert.AnnoDoc = dr.bc_anno
                    monopattinoToInsert.SerieDoc = dr.bc_serie
                    monopattinoToInsert.NumDoc = dr.bc_numdoc
                    monopattinoToInsert.RigaDoc = dr.bc_riga
                    InsertIntoHHOUTMAGDHL(monopattinoToInsert)
                Next
            End If

            Return operazioneOk

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(ElaboraFileRientroTrasferimentiFromB2bTob2c)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "GESTIONE COMUNE DOCUMENTI "

    Private Function initVEBOLL() As Boolean
        Try
            If Not docMagVeboll Is Nothing Then Return True

            docMagVeboll = New CLEVEBOLL
            AddHandler docMagVeboll.RemoteEvent, AddressOf gestisciEventiEntityVeboll

            If Not docMagVeboll.Init(oApp, Nothing, oMenu.oCleComm, "", False, "", "") Then Return False
            If Not docMagVeboll.InitExt() Then Return False
            docMagVeboll.bModuloCRM = False
            docMagVeboll.bIsCRMUser = False
            oMenu = CType(oApp.oMenu, CLE__MENU)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(initVEBOLL)}. Errore: {ex.Message}")
            Return False
        End Try
    End Function

    Private Sub gestisciEventiEntityVeboll(ByVal sender As Object, ByRef e As NTSEventArgs)
        Try

            Dim sMsg As New StringBuilder

            If e.Message.Length <= 0 Then Return

            sMsg.AppendLine("Messaggio dall' Entity VEBOLL:")
            sMsg.AppendLine(e.Message)
            wl(sMsg.ToString())

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(gestisciEventiEntityVeboll)}. Errore: {ex.Message}")
        End Try
    End Sub

    Private Function creaDocNeutroDaRilevazioneDHL(ByVal listaSingoloCarico As List(Of CaricoMagDhl), ByVal doc As String) As Boolean
        Try

            Dim tipoDoc As String = BMI_TIPODOC
            Dim annoDoc As Integer = Date.Today.Year
            Dim serieDoc As String = BMI_SERIE

            'apro il documento
            initVEBOLL()

            Dim lNumTmpDoc As Integer = docMagVeboll.LegNuma(tipoDoc, serieDoc, annoDoc)
            Dim ds As New DataSet
            If Not docMagVeboll.ApriDoc(_dittaCorrente, True, tipoDoc, annoDoc, serieDoc, lNumTmpDoc, ds) Then
                wl("Errore durante l'inizalizzazione del documento.")
                Return False
            End If

            docMagVeboll.bInApriDocSilent = True
            docMagVeboll.ResetVar()
            docMagVeboll.strVisNoteConto = "N"

            If Not docMagVeboll.NuovoDocumento(_dittaCorrente, tipoDoc, annoDoc, serieDoc, lNumTmpDoc) Then
                wl("Errore durante l'inizalizzazione del documento.")
                Return False
            End If

            docMagVeboll.bInNuovoDocSilent = True
            docMagVeboll.bNoDatValDistinta = True

            Dim riferim As String = listaSingoloCarico(0).RifBollaMin

            'Preparazione testata
            With docMagVeboll.dttET.Rows(0)
                !et_conto = BMI_CONTO
                !et_datdoc = DateTime.Today
                !et_tipobf = BMI_TPBF
                !et_magaz = _magazzinoDhl
                !et_riferim = riferim
            End With

            If Not docMagVeboll.OkTestata() Then
                wl("Salvataggio testata non effettuato.")
                Return False
            End If

            'estraggo gli articoli per il documento che sto analizzando
            Dim listaRigheDoc As List(Of Integer)
            listaRigheDoc = listaSingoloCarico.Where(Function(x) x.DocCompleto = doc).Select(Function(y) y.RigaDoc).Distinct.ToList
            For Each numRiga As Integer In listaRigheDoc

                Dim listaElementiRiga As List(Of CaricoMagDhl) = listaSingoloCarico.Where(Function(x) x.DocCompleto = doc And x.RigaDoc = numRiga).ToList
                Dim quant As Integer = listaElementiRiga.Count

                If Not docMagVeboll.AggiungiRigaCorpo(False, listaElementiRiga(0).CodArt, 0, listaElementiRiga(0).RigaDoc) Then
                    wl($"ATTENZIONE. Impossibile inserire la riga nel documento per l'articolo cod: {numRiga}")
                    Return False
                End If

                With docMagVeboll.dttEC.Rows(docMagVeboll.dttEC.Rows.Count - 1)
                    !ec_colli = quant
                    !ec_quant = quant
                End With

                If Not docMagVeboll.RecordSalva(docMagVeboll.dttEC.Rows.Count - 1, False, Nothing) Then
                    docMagVeboll.dttEC.Rows(docMagVeboll.dttEC.Rows.Count - 1).Delete()
                    wl($"Salvataggio riga ordine non effettuata.")
                    Return False
                End If

                For Each riga As CaricoMagDhl In listaElementiRiga
                    aggiungiMatricolaRigaCorpo(tipoDoc, annoDoc, serieDoc, lNumTmpDoc, riga.RigaDoc, riga.RigaMatr, riga.Matricola)
                Next
            Next

            docMagVeboll.CalcolaTotali()
            If Not docMagVeboll.SalvaDocumento("N") Then
                wl($"Impossibile salvare il documento Tipo-Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                Return False
            End If

            wl($"Salvato il documento Tipo-Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}.")

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(creaDocNeutroDaRilevazioneDHL)}. Errore: {ex.Message}")
        Finally
            docMagVeboll.Dispose()
            docMagVeboll = Nothing
        End Try
    End Function

    Public Function CreaBMISpostamentoMerceDaB2bB2c(ByVal tmpObjDocNp As SpedizioniMagDhl, ByRef tmpObjDocBmi As SpedizioniMagDhl,
                                                    ByRef sMsgErr As StringBuilder) As Boolean
        Try

            If tmpObjDocNp Is Nothing Then Return False

            Dim tipoDoc As String = BMI_TIPODOC
            Dim annoDoc As Integer = Date.Today.Year
            Dim serieDoc As String = BMI_SERIE

            initVEBOLL()

            Dim lNumTmpDoc As Integer = docMagVeboll.LegNuma(tipoDoc, serieDoc, annoDoc)

            Dim contoNotaPrelievo As Integer = 0
            Dim tipoBfNotaPrelievo As Integer = 0
            Dim magazNotaprelievo As Integer = 0
            Dim magaz2NotaPrelievo As Integer = 0

            tmpObjDocBmi = New SpedizioniMagDhl()
            tmpObjDocBmi.TipoDoc = tipoDoc
            tmpObjDocBmi.AnnoDoc = annoDoc
            tmpObjDocBmi.SerieDoc = serieDoc
            tmpObjDocBmi.NumDoc = lNumTmpDoc

            'estraggo la nota di prelievo
            Dim listaNotaPrelievo As New List(Of testprb)
            Using dbBaan As New BusinessDbDataContext(_strSqlBusiness)

                listaNotaPrelievo = dbBaan.testprb.Where(Function(x) x.tm_tipork = tmpObjDocNp.TipoDoc And
                                                                     x.tm_anno = tmpObjDocNp.AnnoDoc And
                                                                     x.tm_serie = tmpObjDocNp.SerieDoc And
                                                                     x.tm_numdoc = tmpObjDocNp.NumDoc).ToList()

            End Using

            If listaNotaPrelievo.Count <> 1 Then
                wl($"Nessuna riga trovata per la nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                sMsgErr.AppendLine($"Nessuna riga trovata per la nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                Return False
            End If

            contoNotaPrelievo = listaNotaPrelievo(0).tm_conto
            tipoBfNotaPrelievo = listaNotaPrelievo(0).tm_tipobf
            magazNotaprelievo = listaNotaPrelievo(0).tm_magaz
            magaz2NotaPrelievo = listaNotaPrelievo(0).tm_magaz2
            Dim ds As New DataSet

            If Not docMagVeboll.ApriDoc(_dittaCorrente, True, tipoDoc, annoDoc, serieDoc, lNumTmpDoc, ds) Then
                wl($"Errore in inzializzazione BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                sMsgErr.AppendLine($"Errore in inzializzazione BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                Return False
            End If

            docMagVeboll.bInApriDocSilent = True
            docMagVeboll.ResetVar()
            docMagVeboll.strVisNoteConto = "N"

            If Not docMagVeboll.NuovoDocumento(_dittaCorrente, tipoDoc, annoDoc, serieDoc, lNumTmpDoc) Then
                wl($"Errore in inzializzazione BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                sMsgErr.AppendLine($"Errore in inzializzazione BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                Return False
            End If

            docMagVeboll.bInNuovoDocSilent = True
            docMagVeboll.bNoDatValDistinta = True


            'Preparazione testata
            With docMagVeboll.dttET.Rows(0)
                !et_conto = contoNotaPrelievo
                !et_datdoc = DateTime.Today
                !et_tipobf = tipoBfNotaPrelievo
                !et_magaz = magazNotaprelievo
                !et_magaz2 = magaz2NotaPrelievo
            End With

            If Not docMagVeboll.OkTestata() Then
                wl($"ATTENZIONE Controllo testata BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc} non riuscito.")
                sMsgErr.AppendLine($"ATTENZIONE Controllo testata BMI Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc} non riuscito.")
                Return False
            End If

            Dim dttNotaPrelievo As New DataTable
            Using reader As ObjectReader = ObjectReader.Create(listaNotaPrelievo)

                dttNotaPrelievo.Load(reader)

            End Using

            If dttNotaPrelievo.Rows.Count <> 1 Then
                wl($"Nessuna riga trovata per la nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                sMsgErr.AppendLine($"Nessuna riga trovata per la nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                Return False
            End If

            If Not docMagVeboll.ImportaNote(dttNotaPrelievo) Then
                wl($"Errore in importazione nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                sMsgErr.AppendLine($"Errore in importazione nota di prelievo: Anno-Serie-Numero: {tmpObjDocNp.TipoDoc}-{tmpObjDocNp.AnnoDoc}-{tmpObjDocNp.SerieDoc}-{tmpObjDocNp.NumDoc}")
                Return False
            End If

            docMagVeboll.CalcolaTotali()
            If Not docMagVeboll.SalvaDocumento("N") Then
                wl($"Impossibile salvare il documento Tipo-Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                sMsgErr.AppendLine($"Impossibile salvare il documento Tipo-Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
                Return False
            End If

            wl($"Salvataggio documento Tipo-Anno-Serie-Numero {tipoDoc}-{annoDoc}-{serieDoc}-{lNumTmpDoc}")
            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(CreaBMISpostamentoMerceDaB2bB2c)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function aggiornaDocSpedizioniMagDHLConMatricole(ByVal elementiDocumento As List(Of SpedizioniMagDhl), ByVal doc As String,
                                                             ByRef sMsgErr As StringBuilder) As Boolean

        Try

            Dim operazioniOk As Boolean = True

            'apro il documento
            initVEBOLL()
            Dim dsMyDoc As New DataSet
            If Not docMagVeboll.ApriDoc(_dittaCorrente, False, elementiDocumento(0).TipoDoc, elementiDocumento(0).AnnoDoc, elementiDocumento(0).SerieDoc, elementiDocumento(0).NumDoc, dsMyDoc) Then
                wl("Errore durante l'inizalizzazione del documento.")
                sMsgErr.AppendLine($"Impossibile aprire Nota prelievo Anno-Serie-Numero {elementiDocumento(0).AnnoDoc}-{elementiDocumento(0).SerieDoc}-{elementiDocumento(0).NumDoc}")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            docMagVeboll.bInApriDocSilent = True
            docMagVeboll.ResetVar()
            docMagVeboll.strVisNoteConto = "N"
            docMagVeboll.bCreaFilePick = False

            Dim listaRigheDoc As List(Of Integer)
            listaRigheDoc = elementiDocumento.Where(Function(x) x.DocCompleto = doc).Select(Function(y) y.RigaDoc).Distinct.ToList()

            Dim magazPrelievo As Integer = NTSCInt(docMagVeboll.dttET.Rows(0)!et_magaz)

            For Each rigaDoc As String In listaRigheDoc

                Dim listaElementiRiga As List(Of SpedizioniMagDhl) = elementiDocumento.Where(Function(x) x.DocCompleto = doc And x.RigaDoc = rigaDoc).ToList()
                Dim rigaMatr As Integer = 1

                If docMagVeboll.dttMOVMATR.Rows.Count > 0 Then
                    sMsgErr.AppendLine($"Nota prelievo Anno-Serie-Numero {elementiDocumento(0).AnnoDoc}-{elementiDocumento(0).SerieDoc}-{elementiDocumento(0).NumDoc}")
                    sMsgErr.AppendLine($"Lista matricole già inserita. Impossibile sovrascrivere le matricole nella nota di prelievo.")
                    Return False
                End If

                For Each riga As SpedizioniMagDhl In listaElementiRiga

                    If Not checkMatricolaNotaPrelievo(riga.CodArt, riga.Matricola, magazPrelievo, sMsgErr) Then
                        operazioniOk = False
                        Continue For
                    End If

                    aggiungiMatricolaRigaCorpo(elementiDocumento(0).TipoDoc, elementiDocumento(0).AnnoDoc, elementiDocumento(0).SerieDoc, elementiDocumento(0).NumDoc, riga.RigaDoc, rigaMatr, riga.Matricola)
                    rigaMatr += 1
                Next

            Next

            docMagVeboll.CalcolaTotali()
            If Not docMagVeboll.SalvaDocumento("U") Then
                wl($"Impossibile salvare il documento Tipo-Anno-Serie-Numero {elementiDocumento(0).TipoDoc}-{elementiDocumento(0).AnnoDoc}-{elementiDocumento(0).SerieDoc}-{elementiDocumento(0).NumDoc}")
                sMsgErr.AppendLine($"Impossibile salvare Nota prelievo Anno-Serie-Numero {elementiDocumento(0).AnnoDoc}-{elementiDocumento(0).SerieDoc}-{elementiDocumento(0).NumDoc}")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If operazioniOk Then
                Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                    Dim floNotadhlelab As testprb = dbBusiness.testprb.Where(Function(x) x.tm_tipork = elementiDocumento(0).TipoDoc And
                                                         x.tm_anno = elementiDocumento(0).AnnoDoc And
                                                         x.tm_serie = elementiDocumento(0).SerieDoc And
                                                         x.tm_numdoc = elementiDocumento(0).NumDoc).SingleOrDefault()
                    floNotadhlelab.tm_hhflnotadhlelab = "S"
                    dbBusiness.SubmitChanges()
                End Using
            End If

            wl($"Salvato il documento Tipo-Anno-Serie-Numero {elementiDocumento(0).TipoDoc}-{elementiDocumento(0).AnnoDoc}-{elementiDocumento(0).SerieDoc}-{elementiDocumento(0).NumDoc}")

            Return operazioniOk

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(aggiornaDocSpedizioniMagDHLConMatricole)}. Errore: {ex.Message}")
        Finally
            docMagVeboll.Dispose()
            docMagVeboll = Nothing
        End Try
    End Function

    Private Function checkMatricolaNotaPrelievo(ByVal codArt As String, ByVal matricola As String, ByVal magazPrelievo As Integer,
                                                ByRef sMsgErr As StringBuilder) As Boolean
        Try

            Dim objTmpCheckMatricola As RDM_DHL_StockPf = Nothing
            Dim contValori As Integer = 0

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                contValori = dbBusiness.RDM_DHL_StockPf.Where(Function(x) x.mma_matric = matricola).ToList().Count

            End Using

            If contValori < 1 Then
                sMsgErr.AppendLine($"ATTENZIONE Matricola: {matricola} prelevata non trovata nello stock!")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If contValori > 1 Then
                sMsgErr.AppendLine($"ATTENZIONE Matricola: {matricola} trovata n° {contValori} volte nello stock!")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                objTmpCheckMatricola = dbBusiness.RDM_DHL_StockPf.Where(Function(x) x.mma_matric = matricola).SingleOrDefault()

            End Using

            If objTmpCheckMatricola Is Nothing Then
                sMsgErr.AppendLine($"ATTENZIONE Matricola: {matricola} prelevata non trovata nello stock!")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If objTmpCheckMatricola.Giacenza > 1 Then
                sMsgErr.AppendLine($"ATTENZIONE Matricola: {matricola} ha una giacenza di {objTmpCheckMatricola.Giacenza}.")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If objTmpCheckMatricola.km_codart.Trim() <> codArt.Trim() Then
                sMsgErr.AppendLine($"ATTENZIONE. La matricola: {matricola} è associata all'articolo cod: {objTmpCheckMatricola.km_codart} e non all'articolo cod: {codArt}")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If objTmpCheckMatricola.km_magaz <> magazPrelievo Then
                sMsgErr.AppendLine($"ATTENZIONE. La matricola: {matricola} associata all' articolo cod: {codArt} è nel magazzino {objTmpCheckMatricola.km_magaz} e non nel magazzino della nota di prelievo {magazPrelievo}")
                sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                Return False
            End If

            If Not objTmpCheckMatricola.PrenotatoPerCliente Is Nothing Then
                If objTmpCheckMatricola.PrenotatoPerCliente.Trim() <> String.Empty Then
                    sMsgErr.AppendLine($"ATTENZIONE. La matricola: {matricola} è già stata assegnata al cliente {objTmpCheckMatricola.PrenotatoPerCliente.Trim()}")
                    sMsgErr.AppendLine($"Impossibile complatare inserimento matricole per spedizione DHL.")
                    Return False
                End If
            End If


            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(checkMatricolaNotaPrelievo)}. Errore: {ex.Message}")
        End Try
    End Function

    Private Function aggiungiMatricolaRigaCorpo(ByVal tipork As String, ByVal anno As Integer, ByVal serie As String, ByVal numDoc As Integer,
                                                           ByVal rigaDoc As Integer, ByVal rigaMatr As Integer, ByVal matricola As String) As Boolean
        Try

            'AGGIUNGO MATRICOLA
            docMagVeboll.dttMOVMATR.Rows.Add(docMagVeboll.dttMOVMATR.NewRow)
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!codditt = "."
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!codditt = _dittaCorrente
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_tipork = tipork
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_anno = anno
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_serie = serie
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_numdoc = numDoc
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_riga = rigaDoc
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_rigaa = rigaMatr
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_matric = matricola
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_quant = 1
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1)!mma_hhmotore = " "
            docMagVeboll.dttMOVMATR.Rows(docMagVeboll.dttMOVMATR.Rows.Count - 1).AcceptChanges()

            Return True
        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(aggiungiMatricolaRigaCorpo)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

#End Region

#Region "GESTIONE FILE LOCALI/SFTP"

    Private Function uploadSftpDhlVerona(ByVal nomeCompletoFileLocale As String, ByVal nomeCompletoFileRemoto As String) As Boolean
        Try

            Dim hostSftp As String = String.Empty
            Dim portSftp As Integer = 0
            Dim userSftp As String = String.Empty
            Dim pswSftp As String = String.Empty
            Dim mySftpClient As SftpClient = Nothing

            Dim fileStreamUpload As Stream = Nothing

            'Parametri SFTP
            hostSftp = My.Settings.hostSftp
            portSftp = My.Settings.portSftp
            userSftp = My.Settings.userSftp
            pswSftp = My.Settings.pswSftp

            'Path file

            fileStreamUpload = File.Open(nomeCompletoFileLocale, FileMode.Open)

            Try

                mySftpClient = New SftpClient(hostSftp, portSftp, userSftp, pswSftp)
                mySftpClient.Connect()

            Catch ex As Exception
                wl($"Errore in Metodo {NameOf(GestioneFanticDhl)}. Errore: " & ex.Message)
            End Try

            Try

                mySftpClient.UploadFile(fileStreamUpload, nomeCompletoFileRemoto)

            Catch ex As Exception
                wl($"Errore in Metodo {NameOf(GestioneFanticDhl)}. Errore: " & ex.Message)
            End Try

            mySftpClient.Disconnect()
            mySftpClient.Dispose()
            fileStreamUpload.Close()

            spostaFileElaborato(nomeCompletoFileLocale)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(uploadSftpDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function downloadSftpDhlVerona() As Boolean
        Try

            Dim hostSftp As String = String.Empty
            Dim portSftp As Integer = 0
            Dim userSftp As String = String.Empty
            Dim pswSftp As String = String.Empty
            Dim mySftpClient As SftpClient = Nothing

            Dim outPutSftpDirectory As String = My.Settings.FtpPath
            Dim pathLocale As String = My.Settings.PathCSV
            Dim listaFile As New List(Of SftpFile)
            Dim nomeFileLocale As String = String.Empty
            Dim fileStreamDownload As Stream = Nothing


            'Parametri SFTP
            hostSftp = My.Settings.hostSftp
            portSftp = My.Settings.portSftp
            userSftp = My.Settings.userSftp
            pswSftp = My.Settings.pswSftp

            Try

                mySftpClient = New SftpClient(hostSftp, portSftp, userSftp, pswSftp)
                mySftpClient.Connect()

            Catch ex As Exception
                wl($"Errore in Metodo {NameOf(GestioneFanticDhl)}. Errore: " & ex.Message)
            End Try

            'Scarico CSV Moto
            Try

                outPutSftpDirectory = Path.Combine(outPutSftpDirectory, PATH_USCITA_MOTO_DA_DHL)
                listaFile = mySftpClient.ListDirectory(outPutSftpDirectory)


                For Each myFile As SftpFile In listaFile

                    If Not myFile.IsRegularFile Then
                        Continue For
                    End If

                    nomeFileLocale = Path.Combine(pathLocale, myFile.Name)
                    fileStreamDownload = File.Create(nomeFileLocale)

                    mySftpClient.DownloadFile(myFile.FullName, fileStreamDownload)
                    mySftpClient.DeleteFile(myFile.FullName)
                    fileStreamDownload.Close()

                Next

            Catch ex As Exception
                wl($"Download dei file dal path: {outPutSftpDirectory}. NON RIUSCITA")
                Return False
            End Try

            'Scarico Monopattini 
            Try

                outPutSftpDirectory = My.Settings.FtpPath
                outPutSftpDirectory = Path.Combine(outPutSftpDirectory, PATH_USCITA_MONO_DA_DHL)
                listaFile = mySftpClient.ListDirectory(outPutSftpDirectory)


                For Each myFile As SftpFile In listaFile

                    If Not myFile.IsRegularFile Then
                        Continue For
                    End If

                    nomeFileLocale = Path.Combine(pathLocale, myFile.Name)
                    fileStreamDownload = File.Create(nomeFileLocale)

                    mySftpClient.DownloadFile(myFile.FullName, fileStreamDownload)
                    mySftpClient.DeleteFile(myFile.FullName)
                    fileStreamDownload.Close()

                Next

            Catch ex As Exception
                wl($"Download dei file dal path: {outPutSftpDirectory}. NON RIUSCITA")
                Return False
            End Try

            mySftpClient.Disconnect()
            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(downloadSftpDhlVerona)}. Errore: " & ex.Message)
            Return False
        End Try

    End Function

    Private Sub spostaFileElaborato(ByVal fileSorgenteDaSpostare)
        Try

            Dim pathCsvLocale As String = My.Settings.PathCSV
            Dim nomeFileSpostato As String = String.Empty
            pathCsvLocale = Path.Combine(pathCsvLocale, "Elaborati")

            If Not Directory.Exists(pathCsvLocale) Then
                Directory.CreateDirectory(pathCsvLocale)
            End If


            nomeFileSpostato = Path.Combine(pathCsvLocale, Path.GetFileName(fileSorgenteDaSpostare))

            If File.Exists(nomeFileSpostato) Then
                File.Delete(nomeFileSpostato)
            End If

            File.Move(fileSorgenteDaSpostare, nomeFileSpostato)

            Return

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(spostaFileElaborato)}. Errore: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "METODI GESTIONE BUSINESS"

    Private Function TestCollegaSqlBusiness() As Boolean
        Try

            Dim dbBusiness As BusinessDbDataContext = New BusinessDbDataContext(_strSqlBusiness)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(TestCollegaSqlBusiness)}. Errore: " & ex.Message)
            _flagErroriOk = False
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
            _flagErroriOk = False
            Return False
        End Try
    End Function

    Private Sub GestisciEventiEntityAppBase(sender As Object, ByRef e As NTSEventArgs)
        Try
            If e.Message <> String.Empty Then
                wl(String.Format("Errore interno App. Err: {0}", e.Message))
                _flagErroriOk = False
            End If
        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(GestisciEventiEntityAppBase)} . Errore: " & ex.Message)
        End Try
    End Sub

#End Region

#Region "INVIO MAIL"

    Private Function sendMailErroreMatricoleNotaPrelievo(ByVal sMsgErr As StringBuilder, ByVal tmpObjDoc As SpedizioniMagDhl) As Boolean
        Try

            If tmpObjDoc Is Nothing Then Return False

            Dim par As SMTPParametri = New SMTPParametri
            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                par.SMTPAddress = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPAddress").FirstOrDefault.Valore
                par.SMTPFrom = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPFrom").FirstOrDefault.Valore
                par.SMTPPort = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPPort").FirstOrDefault.Valore
                par.SMTPUser = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPUser").FirstOrDefault.Valore
                par.SMTPPassword = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPPassword").FirstOrDefault.Valore
                par.SMTPCc = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "CCErrSpedDhl").FirstOrDefault.Valore
                par.SMTPTo = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "ToErrSpedDhl").FirstOrDefault.Valore
            End Using

            Dim messaggioPosta As String
            Dim oggettoPosta As String

            oggettoPosta = $"Impossibile complatare inserimento matricole per il documento Tipo:{tmpObjDoc.TipoDoc}, Anno : {tmpObjDoc.AnnoDoc}, Serie : {tmpObjDoc.SerieDoc}, NumDoc: {tmpObjDoc.NumDoc}."
            messaggioPosta = sMsgErr.ToString()

            If Not SMTPSend(messaggioPosta, oggettoPosta, par) Then
                Return False
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(sendMailErroreMatricoleNotaPrelievo)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function sendMailErroreMatricoleBmiSpostamentoB2bB2c(sMsgErr As StringBuilder, tmpObjDoc As SpedizioniMagDhl) As Boolean
        Try

            If tmpObjDoc Is Nothing Then Return False

            Dim par As SMTPParametri = New SMTPParametri
            Using dbBusiness As New BusinessDbDataContext(_strSqlBusiness)

                par.SMTPAddress = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPAddress").FirstOrDefault.Valore
                par.SMTPFrom = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPFrom").FirstOrDefault.Valore
                par.SMTPPort = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPPort").FirstOrDefault.Valore
                par.SMTPUser = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPUser").FirstOrDefault.Valore
                par.SMTPPassword = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "SMTPPassword").FirstOrDefault.Valore
                par.SMTPCc = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "CCErrSpedDhl").FirstOrDefault.Valore
                par.SMTPTo = dbBusiness.HHPARAMETRI.Where(Function(x) x.TipoDato = "ToErrSpedDhl").FirstOrDefault.Valore
            End Using

            Dim messaggioPosta As String
            Dim oggettoPosta As String

            oggettoPosta = $"Impossibile complatare inserimento matricole per il documento Tipo:{tmpObjDoc.TipoDoc}, Anno : {tmpObjDoc.AnnoDoc}, Serie : {tmpObjDoc.SerieDoc}, NumDoc: {tmpObjDoc.NumDoc}."
            messaggioPosta = sMsgErr.ToString()

            If Not SMTPSend(messaggioPosta, oggettoPosta, par) Then
                Return False
            End If

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(sendMailErroreMatricoleBmiSpostamentoB2bB2c)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function

    Public Function SMTPSend(ByVal messaggioPosta As String, ByVal oggettoPosta As String, ByRef par As SMTPParametri) As Boolean
        Try

            Dim smtpServer As New SmtpClient
            Dim mail As New MailMessage
            Dim data As Attachment = Nothing

            smtpServer.Credentials = New Net.NetworkCredential(par.SMTPUser, par.SMTPPassword)
            smtpServer.Port = par.SMTPPort
            smtpServer.Host = par.SMTPAddress

            mail = New MailMessage()
            mail.From = New MailAddress(par.SMTPFrom)

            'Destinatari
            If par.SMTPTo.Contains(";") Then
                Dim addrs As String() = par.SMTPTo.Split(";")
                For Each addr In addrs
                    mail.To.Add(addr)
                Next
            Else
                mail.To.Add(par.SMTPTo)
            End If

            'Conoscenza
            If par.SMTPCc.Trim().Length > 0 Then
                If par.SMTPCc.Contains(";") Then
                    Dim addrs As String() = par.SMTPCc.Split(";")
                    For Each addr In addrs
                        mail.CC.Add(addr)
                    Next
                Else
                    mail.CC.Add(par.SMTPCc)
                End If
            End If

            mail.Subject = oggettoPosta
            mail.Body = messaggioPosta
            smtpServer.Send(mail)

            Return True

        Catch ex As Exception
            wl($"Errore in Metodo {NameOf(SMTPSend)}. Errore: " & ex.Message)
            Return False
        End Try
    End Function


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
        Dim nomeFile As String = _pathLog & "\Gestione_Fantic_Dhl_Log" & DateTime.Now.Day.ToString & "_" & DateTime.Now.Month.ToString & "_" & DateTime.Now.Year & "_" & DateTime.Now.Hour.ToString & "_" & DateTime.Now.Minute.ToString & "_" & DateTime.Now.Second.ToString & ".txt"
        _nomeLogFile = nomeFile
        Try
            Dim f As FileStream = File.Create(_nomeLogFile)
            f.Close()
        Catch ex As Exception
            MessageBox.Show("ATTENZIONE!!! Impossibile creare file di log. Errore: " & ex.Message, "ATTENZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


    End Sub

#End Region

End Class

Public Class SMTPParametri
    Public SMTPAddress As String
    Public SMTPUser As String
    Public SMTPPassword As String
    Public SMTPPort As String
    Public SMTPFrom As String
    Public SMTPTo As String
    Public SMTPCc As String
End Class