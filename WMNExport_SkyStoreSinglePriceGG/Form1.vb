Imports System.IO
Imports System.Windows.Forms

Public Class FormMain
    Public conneSolver As New ADODB.Connection ' eSolver
    Public connWmn As New ADODB.Connection ' WebModaNet

    Public Gruppoarchivi As Object
    Dim flagTGVal(30) As Integer
    Dim j As Integer


    Dim cmdeSolver As ADODB.Command
    Dim rsteS As ADODB.Recordset

    Dim rsteS2 As ADODB.Recordset
    Dim cmdeSolver2 As ADODB.Command

    Dim rsteS3 As ADODB.Recordset
    Dim cmdeSolver3 As ADODB.Command

    Dim cmdeSolverList As ADODB.Command
    Dim rsteSList As New ADODB.Recordset

    Dim cmdeSolverDdtVen As ADODB.Command
    Dim rsteSDdtVen As ADODB.Recordset


    Dim cmdWmn As ADODB.Command
    Dim cmdWmn2 As ADODB.Command
    Dim cmdWmnArt As ADODB.Command
    Dim cmdWmnind As ADODB.Command
    Dim rstWmn As ADODB.Recordset
    Dim rstWmn2 As ADODB.Recordset
    Dim cmdWmnClass As ADODB.Command
    Dim rstWmnClass As ADODB.Recordset

    Dim dataUltAgg As DateTime = DateTime.Now
    Dim str_dataUltAgg As String



    Dim CODARTICOLO_DA = ""
    Dim CODARTICOLO_A = "ZZZZZZZZZZZZZZZZ"
    Dim CODMARCA_DA = ""
    Dim CODMARCA_A = "ZZZZZZZZZZZZZZZZ"
    Dim STAGIONE_DA = ""
    Dim STAGIONE_A = "ZZZZZZZZZZZZZZZZZZZZZZZZZZ"
    Dim LINEA_DA = ""
    Dim LINEA_A = "ZZZZZZZZZZZZZZZZ"
    Dim CODGRUPPO = "XS"
    Dim CODNEG = ""
    Dim DESNEG = ""


    Dim codartDa As String = ""
    Dim codartA As String = ""
    Dim codart As String

    Dim marcaDa As String = ""
    Dim marcaA As String = ""

    Dim LineaDa As String = ""
    Dim LineaA As String = ""

    Dim StagioneDa As String = ""
    Dim StagioneA As String = ""

    Dim listinoDa As String = ""
    Dim listinoA As String = ""

    Dim dtInizioDa As String = ""
    Dim dtInizioA As String = ""

    Dim mydocpath As String = ""
    Dim nameFile As String = ""
    Dim file As System.IO.StreamWriter


    Public versione As String = "4.1"

    Public CODLIS_VEN = ""
    Dim DESLIS_VEN = ""
    Dim DATAVAL_DA = ""
    Dim DATAVAL_A = ""

    Dim CARTELLA_LOG = ""
    Dim OPERATORE = ""


    Private Declare Auto Function GetPrivateProfileString Lib "kernel32" (ByVal lpAppName As String,
        ByVal lpKeyName As String,
        ByVal lpDefault As String,
        ByVal lpReturnedString As System.Text.StringBuilder,
        ByVal nSize As Integer,
        ByVal lpFileName As String) As Integer


    Private Sub LanciaEsportazioneSingola()



        Dim modello As String
        Dim annullato As Integer
        Dim indexPoint As Integer

        Dim dataToday As Date
        Dim str_dataToday As String
        Dim anno As String
        Dim mese As String
        Dim giorno As String
        Dim tabella As String

        Dim exportList As Integer = 0



        dataToday = Date.Today.ToString("dd MM yyyy")
        anno = dataToday.Year
        mese = dataToday.Month
        giorno = dataToday.Day

        str_dataUltAgg = dataUltAgg.ToString("dd/MM/yyyy HH:mm:ss")
        str_dataUltAgg = Replace(str_dataUltAgg, ".", ":")

        'str_dataToday = "29/11/2017" ' data solo per db di test
        str_dataToday = " CONVERT(date,'" & dataToday & "', 105) "

        Dim dataFine As String = " CONVERT(date,'31/12/2099', 105) "


        cmdeSolver = New ADODB.Command
        cmdeSolver.ActiveConnection = conneSolver ' .ConnectionString

        cmdeSolver2 = New ADODB.Command
        cmdeSolver2.ActiveConnection = conneSolver

        cmdeSolver3 = New ADODB.Command
        cmdeSolver3.ActiveConnection = conneSolver

        cmdeSolverList = New ADODB.Command
        cmdeSolverList.ActiveConnection = conneSolver ' .ConnectionString

        cmdeSolverDdtVen = New ADODB.Command
        cmdeSolverDdtVen.ActiveConnection = conneSolver ' .ConnectionString

        cmdWmn = New ADODB.Command
        cmdWmn.let_ActiveConnection(connWmn)

        cmdWmn2 = New ADODB.Command
        cmdWmn2.let_ActiveConnection(connWmn)

        cmdWmnArt = New ADODB.Command
        cmdWmnArt.let_ActiveConnection(connWmn)

        cmdWmnind = New ADODB.Command
        cmdWmnind.let_ActiveConnection(connWmn)

        cmdWmnClass = New ADODB.Command
        cmdWmnClass.let_ActiveConnection(connWmn)

        Application.DoEvents()


        ' se selezionato flag codice articolo da .. a ..
        If CODARTICOLO_DA <> "" Or CODARTICOLO_A <> "" Then

            codartDa = CODARTICOLO_DA
            codartA = CODARTICOLO_A

            ' controllo filtro linea
            If LINEA_DA <> "" Or LINEA_A <> "" Then
                LineaDa = LINEA_DA
                LineaA = LINEA_A
            End If

            '' controllo filtro marca
            If CODMARCA_DA <> "" Or CODMARCA_A <> "" Then
                marcaDa = CODMARCA_DA
                marcaA = CODMARCA_A
            End If

            '' controllo filtro stagione
            If STAGIONE_DA <> "" Or STAGIONE_A <> "" Then
                StagioneDa = STAGIONE_DA
                StagioneA = STAGIONE_A
            End If

            'tabella = "DocUniRigheDdtVen"

            Call SelectArticoli(tabella, codartDa, codartA, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            Do Until rsteS.EOF  ' ciclo su tutti gli articoli selezionati
                codart = rsteS.Fields("CodArt").Value

                Label2.Text = codart

                Call SelectDetailArticoli(codart, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

                indexPoint = codart.IndexOf(".")
                modello = Strings.Left(codart, indexPoint)

                If rsteS2.Fields("StatoArt").Value = 1 Then
                    annullato = 1
                Else
                    annullato = 0
                End If

                Call insertTieArticoli(codart, modello, annullato, str_dataToday)

                If rsteS2.Fields("CodiceTabellaTaglie").Value = "" Then
                    Call CompilaValiditaTG("UN", rsteS2)
                Else
                    Call CompilaValiditaTG(rsteS2.Fields("CodiceTabellaTaglie").Value, rsteS2)
                End If

                If CODLIS_VEN = "" Then
                    exportList = 1

                    listinoDa = ""
                    listinoA = ""

                    dtInizioDa = ""
                    dtInizioA = ""
                Else
                    exportList = 1
                    listinoDa = CODLIS_VEN
                    listinoA = CODLIS_VEN

                    If DATAVAL_DA <> "" Or DATAVAL_A <> "" Then
                        dtInizioDa = DATAVAL_DA
                        dtInizioA = DATAVAL_A
                    End If

                End If

                If rsteS2.Fields("CodiceTabellaTaglie").Value = "" Then
                    Call insertTieArtVarianti("UN", codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)
                Else
                    Call insertTieArtVarianti(rsteS2.Fields("CodiceTabellaTaglie").Value, codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)
                End If

                ' Call insertTieArtVarianti(codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)

                rsteS.MoveNext()
                Application.DoEvents()
            Loop

            'tabella = "DocUniRigheFattVen"  ' cambio nome tabella e rifaccio tutto il giro anche per le fatture

            'Call SelectArticoli(tabella, codartDa, codartA, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            'Do Until rsteS.EOF
            '    codart = rsteS.Fields("CodArt").Value

            '    Label2.Text = codart

            '    Call SelectDetailArticoli(codart, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            '    indexPoint = codart.IndexOf(".")
            '    modello = Strings.Left(codart, indexPoint)

            '    If rsteS2.Fields("StatoArt").Value = 1 Then
            '        annullato = 1
            '    Else
            '        annullato = 0
            '    End If

            '    Call insertTieArticoli(codart, modello, annullato, str_dataToday)

            '    Call CompilaValiditaTG(rsteS2.Fields("CodiceTabellaTaglie").Value, rsteS2)

            '    If CODLIS_VEN = "" Then
            '        exportList = 1

            '        listinoDa = ""
            '        listinoA = ""

            '        dtInizioDa = ""
            '        dtInizioA = ""
            '    Else
            '        exportList = 1
            '        listinoDa = CODLIS_VEN
            '        listinoA = CODLIS_VEN

            '        If DATAVAL_DA <> "" Or DATAVAL_A <> "" Then
            '            dtInizioDa = DATAVAL_DA
            '            dtInizioA = DATAVAL_A
            '        End If

            '    End If

            '    Call insertTieArtVarianti(codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)

            '    rsteS.MoveNext()
            '    Application.DoEvents()
            'Loop


        ElseIf (CODARTICOLO_DA = "" And CODARTICOLO_A = "") And ((LINEA_DA <> "" And LineaA <> "") Or (CODMARCA_DA <> "" And CODMARCA_A <> "") Or (STAGIONE_DA <> "" And STAGIONE_A <> "")) Then


            If LINEA_DA <> "" And LineaA <> "" Then
                LineaDa = LINEA_DA
                LineaA = LINEA_A
            End If

            ' controllo filtro marca
            If CODMARCA_DA <> "" And CODMARCA_A <> "" Then
                marcaDa = CODMARCA_DA
                marcaA = CODMARCA_A
            End If

            ' controllo filtro stagione
            If STAGIONE_DA <> "" And STAGIONE_A <> "" Then
                StagioneDa = STAGIONE_DA
                StagioneA = STAGIONE_A
            End If

            ' tabella = "DocUniRigheDdtVen"

            Call SelectArticoli(tabella, codartDa, codartA, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            Do Until rsteS.EOF
                codart = rsteS.Fields("CodArt").Value

                Label2.Text = codart

                Call SelectDetailArticoli(codart, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

                indexPoint = codart.IndexOf(".")
                modello = Strings.Left(codart, indexPoint)

                If rsteS2.Fields("StatoArt").Value = 1 Then
                    annullato = 1
                Else
                    annullato = 0
                End If

                Call insertTieArticoli(codart, modello, annullato, str_dataToday)

                If rsteS2.Fields("CodiceTabellaTaglie").Value = "" Then
                    Call CompilaValiditaTG("UN", rsteS2)
                Else
                    Call CompilaValiditaTG(rsteS2.Fields("CodiceTabellaTaglie").Value, rsteS2)
                End If

                'Call CompilaValiditaTG(rsteS2.Fields("CodiceTabellaTaglie").Value, rsteS2)

                If CODLIS_VEN = "" Then
                    exportList = 1

                    listinoDa = ""
                    listinoA = ""

                    dtInizioDa = ""
                    dtInizioA = ""
                Else
                    exportList = 1
                    listinoDa = CODLIS_VEN
                    listinoA = CODLIS_VEN

                    If DATAVAL_DA <> "" Or DATAVAL_A <> "" Then
                        dtInizioDa = DATAVAL_DA
                        dtInizioA = DATAVAL_A
                    End If

                End If

                If rsteS2.Fields("CodiceTabellaTaglie").Value = "" Then
                    Call insertTieArtVarianti("UN", codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)
                Else
                    Call insertTieArtVarianti(rsteS2.Fields("CodiceTabellaTaglie").Value, codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)
                End If

                'Call insertTieArtVarianti(codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)

                rsteS.MoveNext()
                Application.DoEvents()
            Loop

            'tabella = "DocUniRigheFattVen"

            'Call SelectArticoli(tabella, codartDa, codartA, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            'Do Until rsteS.EOF
            '    codart = rsteS.Fields("CodArt").Value

            '    Label2.Text = codart

            '    Call SelectDetailArticoli(codart, LineaDa, LineaA, marcaDa, marcaA, StagioneDa, StagioneA)

            '    indexPoint = codart.IndexOf(".")
            '    modello = Strings.Left(codart, indexPoint)

            '    If rsteS2.Fields("StatoArt").Value = 1 Then
            '        annullato = 1
            '    Else
            '        annullato = 0
            '    End If

            '    Call insertTieArticoli(codart, modello, annullato, str_dataToday)

            '    Call CompilaValiditaTG(rsteS2.Fields("CodiceTabellaTaglie").Value, rsteS2)

            '    If CODLIS_VEN = "" Then
            '        exportList = 1

            '        listinoDa = ""
            '        listinoA = ""

            '        dtInizioDa = ""
            '        dtInizioA = ""
            '    Else
            '        exportList = 1
            '        listinoDa = CODLIS_VEN
            '        listinoA = CODLIS_VEN

            '        If DATAVAL_DA <> "" Or DATAVAL_A <> "" Then
            '            dtInizioDa = DATAVAL_DA
            '            dtInizioA = DATAVAL_A
            '        End If

            '    End If

            '    Call insertTieArtVarianti(codart, annullato, str_dataToday, dataFine, exportList, listinoDa, listinoA, dtInizioDa, dtInizioA)

            '    rsteS.MoveNext()
            '    Application.DoEvents()
            'Loop

        End If



        rsteS.Close()

        cmdeSolver = Nothing
        cmdeSolver = New ADODB.Command
        cmdeSolver.ActiveConnection = conneSolver ' .ConnectionString


        Application.DoEvents()

        MsgBox("ESPORTAZIONE TERMINATA")



        '   rstWmn2 = Nothing
        '   cmdWmn = Nothing


    End Sub

    Private Sub readIniFile(filename As String)
        Dim VL_FileName As String = filename
        Dim sb As System.Text.StringBuilder
        Dim Sezione As String = "FILTRI"
        Dim Sezione_Server As String = "SERVER"
        Dim i As Integer


        '-----------------------------------------------------------------
        ' FILTRI
        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODARTICOLO_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODARTICOLO_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODARTICOLO_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODARTICOLO_A = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODMARCA_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODMARCA_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODMARCA_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODMARCA_A = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "STAGIONE_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            STAGIONE_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "STAGIONE_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            STAGIONE_A = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "LINEA_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            LINEA_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "LINEA_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            LINEA_A = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODGRUPPO", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODGRUPPO = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODLIS_VEN", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODLIS_VEN = sb.ToString
        End If


        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DESLIS_VEN", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DESLIS_VEN = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "CODNEG", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CODNEG = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DESNEG", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DESNEG = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DATAVAL_DA", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DATAVAL_DA = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "DATAVAL_A", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            DATAVAL_A = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "PATH_LOG", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            CARTELLA_LOG = sb.ToString
        End If

        sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Sezione, "OPERATORE", "", sb, sb.Capacity, VL_FileName)
        If sb.ToString <> "" Then
            OPERATORE = sb.ToString
        End If

    End Sub

    Private Sub connectToDB()
        Dim ServereS As String ' ==> Server eSolver
        Dim UsereS As String ' ==> User eSolver
        Dim PwdeS As String ' ==> Password eSolver
        Dim CateS As String ' ==> Catalogo eSolver

        Dim ServerWmn As String ' ==> Server WebModaNet
        Dim UserWmn As String ' ==> User WebModaNet
        Dim PwdWmn As String ' ==> Password WebModaNet
        Dim CatWmn As String ' ==> Catalogo WebModaNet

        Dim errore_connessionees As Short ' ==> Errore in fase di connessione al Db eSolver
        Dim errore_connessioneWmn As Short ' ==> Errore in fase di connessione al Db WebModaNet

        Dim sConnstring As Object

        Dim pat As Object
        Dim rs As New ADODB.Recordset


        Application.DoEvents()

        ' ==> lettura file ini
        Dim FileName As String = My.Application.Info.DirectoryPath & "\ordiniweb.ini"
        Dim ApplicationKey As String = "Sezione1"

        ' ==> Inizializzazione variabili da file ini

        ' ==> Inizializzazione variabili connessioni ODBC
        Dim sb As System.Text.StringBuilder



        sb = New System.Text.StringBuilder(500)

        UserWmn = ""
        GetPrivateProfileString(ApplicationKey, "userweb", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            UserWmn = sb.ToString
        End If



        PwdWmn = ""
        GetPrivateProfileString(ApplicationKey, "pwdweb", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            PwdWmn = sb.ToString
        End If


        CatWmn = ""
        GetPrivateProfileString(ApplicationKey, "catweb", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            CatWmn = sb.ToString
        End If

        ServerWmn = ""
        GetPrivateProfileString(ApplicationKey, "dsweb", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            ServerWmn = sb.ToString
        End If



        UsereS = ""
        GetPrivateProfileString(ApplicationKey, "usertsm", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            UsereS = sb.ToString
        End If

        PwdeS = ""
        GetPrivateProfileString(ApplicationKey, "pwdtsm", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            PwdeS = sb.ToString
        End If

        CateS = ""
        GetPrivateProfileString(ApplicationKey, "cattsm", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            CateS = sb.ToString
        End If

        ServereS = ""
        GetPrivateProfileString(ApplicationKey, "dstsm", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            ServereS = sb.ToString
        End If



        Gruppoarchivi = ""
        GetPrivateProfileString(ApplicationKey, "gruppo", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            Gruppoarchivi = sb.ToString
        End If


        pat = ""
        GetPrivateProfileString(ApplicationKey, "percorsoesolver", "", sb, sb.Capacity, FileName)
        If sb.ToString <> "" Then
            pat = sb.ToString
        End If


        errore_connessionees = 0
        sConnstring = "Provider=SQLOLEDB;user id = '" & UsereS & "'; password = '" & PwdeS & "' ; initial catalog= '" & CateS & "' ; data source = '" & ServereS & "'"
        conneSolver.CommandTimeout = 600
        Try
            conneSolver.Open(sConnstring)

        Catch ex As Exception
            errore_connessionees = 1
            MsgBox("Errore di connessione a SQL server di Esolver " & ex.Message)


        End Try

        sConnstring = "Provider=SQLOLEDB;user id = '" & UserWmn & "'; password = '" & PwdWmn & "' ; initial catalog= '" & CatWmn & "' ; data source = '" & ServerWmn & "'"

        errore_connessioneWmn = 0
        connWmn.CommandTimeout = 600

        Try
            connWmn.Open(sConnstring)

        Catch ex As Exception
            errore_connessioneWmn = 1
            MsgBox("Errore di connessione a SQL server di Esolver " & ex.Message)

        End Try


    End Sub

    Private Sub SelectArticoli(tabella As String, codArtDa As String, codArtA As String, LineaDa As String, LineaA As String, marcaDa As String, marcaA As String, StagioneDa As String, StagioneA As String)
        Dim sql As String = ""
        sql = "SELECT * "
        sql = sql + " from ArtAnagrafica INNER JOIN ModaArticoli ON ArtAnagrafica.CodArt = ModaArticoli.CodiceArticolo and ArtAnagrafica.DBGruppo=ModaArticoli.DBGruppo "
        sql = sql + " WHERE (ArtAnagrafica.CodGruppoCodifica = 'GGCAL' OR ArtAnagrafica.CodGruppoCodifica = 'GGART') and ArtAnagrafica.CodFamiglia <> 'CTL' and ArtAnagrafica.DbGruppo = 'GO' and ArtAnagrafica.codart LIKE '%.%' and (ArtAnagrafica.StatoArt<>'1' and ArtAnagrafica.StatoArt <> '2')"

        If codArtDa <> "" And codArtA <> "" Then
            sql = sql + " and codart >='" & codArtDa & "' and codart<='" & codArtA & "' "
        Else
            sql = sql + " and codart <> '' "
        End If

        If LineaDa <> "" And LineaA <> "" Then
            sql = sql + " and ModaArticoli.CodLinea >= '" & LineaDa & "' and ModaArticoli.CodLinea <= '" & LineaA & "' "
        End If

        If marcaDa <> "" And marcaA <> "" Then
            sql = sql + " and ArtAnagrafica.CodMarca >= '" & marcaDa & "' and ArtAnagrafica.CodMarca <= '" & marcaA & "' "
        End If

        If StagioneDa <> "" And StagioneA <> "" Then
            sql = sql + " and ModaArticoli.CodStagione >= '" & StagioneDa & "' and ModaArticoli.CodStagione <= '" & StagioneA & "' "
        End If


        cmdeSolver.CommandText = sql
        cmdeSolver.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmdeSolver.Execute()

        rsteS = New ADODB.Recordset
        rsteS.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rsteS.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rsteS.Open(cmdeSolver)

    End Sub

    Private Sub SelectDetailArticoli(codart As String, LineaDa As String, LineaA As String, marcaDa As String, marcaA As String, StagioneDa As String, StagioneA As String)
        Dim Sql As String

        Sql = "SELECT TOP 1 DesArt, DesEstesa, MagUM, VenUM, AcqUM, CodMarca, StatoArt, ModaArticoli.CodiceTabellaTaglie, ModaArticoli.CodStagione, ModaArticoli.CodStatistico3,"
        Sql = Sql + " ModaArticoli.FlagTagliaValida_1,ModaArticoli.FlagTagliaValida_2,ModaArticoli.FlagTagliaValida_3,ModaArticoli.FlagTagliaValida_4,ModaArticoli.FlagTagliaValida_5,ModaArticoli.FlagTagliaValida_6,ModaArticoli.FlagTagliaValida_7,ModaArticoli.FlagTagliaValida_8,ModaArticoli.FlagTagliaValida_9,ModaArticoli.FlagTagliaValida_10,"
        Sql = Sql + " ModaArticoli.FlagTagliaValida_11,ModaArticoli.FlagTagliaValida_12,ModaArticoli.FlagTagliaValida_13,ModaArticoli.FlagTagliaValida_14,ModaArticoli.FlagTagliaValida_15,ModaArticoli.FlagTagliaValida_16,ModaArticoli.FlagTagliaValida_17,ModaArticoli.FlagTagliaValida_18,ModaArticoli.FlagTagliaValida_19,ModaArticoli.FlagTagliaValida_20,"
        Sql = Sql + " ModaArticoli.FlagTagliaValida_21, ModaArticoli.FlagTagliaValida_22,ModaArticoli.FlagTagliaValida_23,ModaArticoli.FlagTagliaValida_24,ModaArticoli.FlagTagliaValida_25,ModaArticoli.FlagTagliaValida_26,ModaArticoli.FlagTagliaValida_27,ModaArticoli.FlagTagliaValida_28,ModaArticoli.FlagTagliaValida_29,ModaArticoli.FlagTagliaValida_30, ArtAnagrafica.CodFamiglia "
        Sql = Sql & " FROM ArtAnagrafica INNER JOIN ModaArticoli ON ArtAnagrafica.CodArt = ModaArticoli.CodiceArticolo"
        Sql = Sql & " WHERE ArtAnagrafica.CodArt = '" & codart & "' "



        cmdeSolver2.CommandText = Sql
        cmdeSolver2.CommandType = ADODB.CommandTypeEnum.adCmdText
        Try
            cmdeSolver2.Execute()
        Catch ex As Exception
            'MsgBox("SELECT DETAIL ARTICOLI")
            'MsgBox(ex.Message)
            'MsgBox(cmdeSolver2.CommandText)
        End Try

        rsteS2 = New ADODB.Recordset
        rsteS2.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rsteS2.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rsteS2.Open(cmdeSolver2)

    End Sub

    Private Sub insertTieArticoli(codart As String, modello As String, annullato As Integer, str_dataToday As String)

        Dim desart As String = rsteS2.Fields("DesArt").Value

        If desart.Length > 0 Then
            desart = Replace(rsteS2.Fields("DesArt").Value, "'", "")
        Else
            desart = ""
        End If

        If desart <> "" And desart.Length > 50 Then
            desart = Strings.Left(desart, 50)
        End If

        Dim desartEst As String = rsteS2.Fields("DesEstesa").Value

        If desartEst.Length > 0 Then
            desartEst = Replace(rsteS2.Fields("DesEstesa").Value, "'", "")
        Else
            desartEst = ""
        End If

        If desartEst <> "" And desartEst.Length > 70 Then
            desartEst = Strings.Left(desartEst, 70)
        End If

        cmdWmn2.CommandText = "SELECT * FROM tieArticoli WHERE CodPadre = '" & codart & "'"
        cmdWmn2.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmdWmn2.Execute()

        rstWmn2 = New ADODB.Recordset
        rstWmn2.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rstWmn2.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rstWmn2.Open(cmdWmn2)

        If rstWmn2.RecordCount = 0 Then
            cmdWmn.CommandText = "INSERT INTO tieArticoli VALUES ('" & codart & "', '" & desart & "', '" & desartEst & "', '" & rsteS2.Fields("MagUM").Value & "', '" & rsteS2.Fields("VenUM").Value & "', '" & rsteS2.Fields("AcqUM").Value & "', 1, 1, 0, '23', '" & rsteS2.Fields("CodMarca").Value & "', '" & rsteS2.Fields("CodiceTabellaTaglie").Value & "', '', '" & rsteS2.Fields("CodStagione").Value & "', '', '" & modello & "', 0, 0, '', " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0, '', '" & rsteS2.Fields("CodStatistico3").Value & "')"
            cmdWmn.CommandType = ADODB.CommandTypeEnum.adCmdText
            Try
                cmdWmn.Execute()
            Catch ex As Exception
                'MsgBox("INSERT TIE ARTICOLI")
                'MsgBox(ex.Message)
                'MsgBox(cmdWmn.CommandText)
            End Try


            cmdWmnClass.CommandText = "SELECT * FROM tieArtClass WHERE Cod_PK = '" & codart & "_" & rsteS2.Fields("CodFamiglia").Value & "'"
            cmdWmnClass.CommandType = ADODB.CommandTypeEnum.adCmdText
            cmdWmnClass.Execute()

            rstWmnClass = New ADODB.Recordset
            rstWmnClass.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
            rstWmnClass.LockType = ADODB.LockTypeEnum.adLockOptimistic
            rstWmnClass.Open(cmdWmnClass)

            If rstWmnClass.RecordCount = 0 Then
                ' inserimento riga nel tieArtClass se non presente
                cmdWmnArt.CommandText = "INSERT INTO tieArtClass VALUES ('" & codart & "_" & rsteS2.Fields("CodFamiglia").Value & "', '" & rsteS2.Fields("CodFamiglia").Value & "', '" & codart & "', " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0, '')"
                cmdWmnArt.CommandType = ADODB.CommandTypeEnum.adCmdText
                cmdWmnArt.Execute()
            End If

            rstWmnClass.Close()


            '' inserimento riga nel tieArtClass
            'cmdWmnArt.CommandText = "INSERT INTO tieArtClass VALUES ('" & codart & "_" & rsteS2.Fields("CodFamiglia").Value & "', '" & rsteS2.Fields("CodFamiglia").Value & "', '" & codart & "', " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0, '')"
            'cmdWmnArt.CommandType = ADODB.CommandTypeEnum.adCmdText
            'cmdWmnArt.Execute()

        End If

    End Sub

    Private Sub CompilaValiditaTG(codtabltaglie As String, rsArt As ADODB.Recordset)
        Dim rs As New ADODB.Recordset
        Dim i As Integer
        ' azzera tutte le validità taglie
        For i = 1 To 30
            flagTGVal(i) = 0
        Next

        If codtabltaglie <> "UN" Then
            ' compila a 1 solo quelle da valorizzare
            rs.Open("select * from modatabellataglie  where DBGruppo = '" & Gruppoarchivi & "' and codicetabellataglie='" & codtabltaglie & "'", conneSolver, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockReadOnly)
            If rs.RecordCount = 1 Then
                For i = 1 To 30
                    If (Trim(rs.Fields("codicitaglie_" & i).Value) <> "") And (rsArt.Fields("FlagTagliaValida_" & i).Value = 1) Then
                        flagTGVal(i) = 1
                    End If
                Next
            End If
            rs.Close()
        End If

        If codtabltaglie = "UN" Then
            flagTGVal(2) = 1
        End If

    End Sub

    Private Sub insertTieArtVarianti(codtabtaglie As String, codart As String, annullato As Integer, str_dataToday As String, dataFine As String, exportList As Integer, listDa As String, listA As String, dtInizVal As String, dtFineVal As String)
        Dim codice As String
        Dim taglia As String

        Dim classeListino As Integer
        Dim chiaveList As String

        Dim inizio As DateTime
        Dim inizio_str As String
        Dim fine_str As String

        '  If exportList = 0 Then
        cmdWmn2.CommandText = "SELECT * FROM tieArtVarianti WHERE CodPadre = '" & codart & "'"
        cmdWmn2.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmdWmn2.Execute()

        rstWmn2 = New ADODB.Recordset
        rstWmn2.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rstWmn2.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rstWmn2.Open(cmdWmn2)

        If rstWmn2.RecordCount = 0 Then
            For i = 1 To 30
                If flagTGVal(i) = 1 Then
                    If codtabtaglie <> "UN" Then
                        codice = codart + i.ToString("00")
                    Else
                        codice = codart
                    End If

                    If codice.Contains("KIT") = False And codtabtaglie = "UN" Then
                        codice = codice + i.ToString("00")
                    End If

                    taglia = codtabtaglie & "_" & i.ToString()

                    cmdWmn.CommandText = "INSERT INTO tieArtVarianti VALUES ('" & codice & "', '" & codart & "', '', '" & taglia & "', 1, 1, " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0, '')"
                    cmdWmn.CommandType = ADODB.CommandTypeEnum.adCmdText
                    Try

                        cmdWmn.Execute()
                    Catch ex As Exception
                        'MsgBox("INSERT TIE VARIANTI")
                        'MsgBox(ex.Message)
                        'MsgBox(cmdWmn.CommandText)
                    End Try
                End If
            Next
        End If
        ' End If

        If exportList = 1 Then

            classeListino = 1

            Call SelectPriceList(codart, classeListino, listDa, listA, dtInizVal, dtFineVal)



            file = My.Computer.FileSystem.OpenTextFileWriter(CARTELLA_LOG & "\" & nameFile, True)

            file.WriteLine("---" & OPERATORE & "---" & DateTime.Now.ToString() & "----- INIZIO ELABORAZIONE -------------------" & vbCrLf)
            file.WriteLine("Codice Articolo da " & CODARTICOLO_DA & " a " & CODARTICOLO_A & vbCrLf)
            file.WriteLine("Codice marca da " & CODMARCA_DA & " a " & CODMARCA_A & vbCrLf)
            file.WriteLine("Codice linea da " & LINEA_DA & " a " & LINEA_A & vbCrLf)
            file.WriteLine("Codice stagione da " & STAGIONE_DA & " a " & STAGIONE_A & vbCrLf)

            Do Until rsteSList.EOF
                fine_str = "20991231"
                inizio = rsteSList.Fields("DataInizioValidita").Value
                inizio_str = inizio.Year & inizio.Month.ToString("00") & inizio.Day.ToString("00")

                'chiaveList = codart & "\" & rsteSList.Fields("CodiceNegozioSky").Value & "\" & rsteSList.Fields("CodValuta").Value & "\" & inizio_str & "\" & fine_str

                Call insertTiePrzVendita(codtabtaglie, chiaveList, codart, codice, annullato, str_dataToday, classeListino, dataFine, fine_str, inizio_str)

                rsteSList.MoveNext()
                Application.DoEvents()
            Loop
            file.WriteLine("-----------------------" & DateTime.Now.ToString() & "----- FINE ELABORAZIONE -------------------" & vbCrLf)
            file.Close()

            ' 03/05/2018 - commentato per il momento scrittura listini di acquisto
            '
            'classeListino = 2
            'Call SelectPriceList(codart, classeListino, listDa, listA, dtInizVal, dtFineVal)



            'Do Until rsteSList.EOF
            '    fine_str = "20991231"
            '    inizio = rsteSList.Fields("DataInizioValidita").Value
            '    inizio_str = inizio.Year & inizio.Month.ToString("00") & inizio.Day.ToString("00")

            '    ' chiaveList = codart & "\" & rsteSList.Fields("CodiceNegozioSky").Value & "\" & rsteSList.Fields("CodValuta").Value & "\" & inizio_str & "\" & fine_str

            '    Call insertTiePrzVendita(chiaveList, codart, codice, annullato, str_dataToday, classeListino, dataFine, fine_str, inizio_str)

            '    rsteSList.MoveNext()
            '    Application.DoEvents()
            'Loop

        End If
    End Sub

    Private Sub SelectPriceList(codart As String, classeListino As Integer, listDa As String, listA As String, dtInizVal As String, dtFineVal As String)
        Dim sql As String = ""

        sql = "SELECT * FROM LISTINI_RIGHE INNER JOIN Listini ON LISTINI_RIGHE.CodListino = Listini.CodListino "
        sql = sql + " INNER JOIN GG_NEGOZI ON (LISTINI_RIGHE.CodListino = GG_NEGOZI.ListinoVendite) " ' OR LISTINI_RIGHE.CodListino = GG_NEGOZI.ListinoAcquisti) "
        sql = sql + " WHERE LISTINI_RIGHE.DbGruppo = '" & Gruppoarchivi & "' and LISTINI_RIGHE.ClasseListino = '" & classeListino.ToString() & "' "

        If listDa <> "" And listA <> "" Then
            sql = sql + " and LISTINI_RIGHE.CodListino >= '" & listDa & "' and LISTINI_RIGHE.CodListino <= '" & listA & "' "
        Else
            sql = sql + " and LISTINI_RIGHE.CodListino <> '' "
        End If

        If CODNEG <> "" Then
            sql = sql + " and GG_NEGOZI.CodiceNegozioSky = '" & CODNEG & "' "
            'Else
            '    sql = sql + " and GG_NEGOZI.CodiceNegozioSky <> '' "
        End If

        sql = sql + " and CodArt = '" & codart & "' "

        If dtInizVal <> "" And dtFineVal <> "" Then
            sql = sql + " and DataInizioValidita >= '" & dtInizVal & "' and DataInizioValidita <= '" & dtFineVal & "'"
        End If

        'Clipboard.SetText(sql)

        cmdeSolverList.CommandText = sql
        cmdeSolverList.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmdeSolverList.Execute()

        rsteSList = New ADODB.Recordset
        rsteSList.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rsteSList.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rsteSList.Open(cmdeSolverList)


    End Sub

    Private Sub insertTiePrzVendita(codtabtaglie As String, chiaveList As String, codart As String, codice As String, annullato As Integer, str_dataToday As String, classeListino As Integer, dataFine As String, fine_str As String, inizio_str As String)
        Dim prezzo_str As String = "" 'Replace(rsteSList.Fields("Prezzo").Value.ToString(), ",", ".")
        chiaveList = codice & "\" & rsteSList.Fields("CodiceNegozioSky").Value & "\" & rsteSList.Fields("CodValuta").Value & "\" & inizio_str & "\" & fine_str

        If classeListino = 1 Then
            cmdWmn2.CommandText = "SELECT * FROM tiePrzVendita WHERE cod_PK = '" & chiaveList & "'"
        Else
            cmdWmn2.CommandText = "SELECT * FROM tiePrzAcquisto WHERE cod_PK = '" & chiaveList & "'"
        End If
        cmdWmn2.CommandType = ADODB.CommandTypeEnum.adCmdText
        cmdWmn2.Execute()

        rstWmn2 = New ADODB.Recordset
        rstWmn2.CursorType = ADODB.CursorTypeEnum.adOpenKeyset
        rstWmn2.LockType = ADODB.LockTypeEnum.adLockOptimistic
        rstWmn2.Open(cmdWmn2)



        If rstWmn2.RecordCount = 0 Then
            If classeListino = 1 Then
                For i = 1 To 30
                    If flagTGVal(i) = 1 Then
                        If codtabtaglie <> "UN" Then
                            codice = codart & i.ToString("00")
                        Else
                            codice = codart
                        End If

                        If codice.Contains("KIT") = False And codtabtaglie = "UN" Then
                            codice = codice + i.ToString("00")
                        End If

                        prezzo_str = Replace(rsteSList.Fields("PrezzoTaglia" & i.ToString()).Value.ToString(), ",", ".")
                        chiaveList = codice & "\" & rsteSList.Fields("CodiceNegozioSky").Value & "\" & rsteSList.Fields("CodValuta").Value & "\" & inizio_str & "\" & fine_str
                        Try
                            'metto codart al posto di codice (codart e poi codice)
                            cmdWmn.CommandText = "INSERT INTO tiePrzVendita VALUES ('" & chiaveList & "', '" & codart & "', '" & codice & "', '" & rsteSList.Fields("CodiceNegozioSky").Value & "', '',1, '" & rsteSList.Fields("CodValuta").Value & "', " & prezzo_str & ",0, '" & rsteSList.Fields("DataInizioValidita").Value & "', " & dataFine & ", " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0 , '', '')"
                            cmdWmn.CommandType = ADODB.CommandTypeEnum.adCmdText
                            cmdWmn.Execute()

                            'file.WriteLine("Operatore e filtri | " & OPERATORE & ";" & CODARTICOLO_DA & ";" & CODARTICOLO_A & ";" & CODMARCA_DA & ";" & CODMARCA_A & ";" & STAGIONE_DA & ";" & STAGIONE_A & ";" & LINEA_DA & ";" & LINEA_A & ";" & CODLIS_VEN & ";" & vbCrLf)
                            'file.WriteLine("Valori esportati | " & rsteSList.Fields("CodiceNegozioSky").Value & ";" & chiaveList & ";" & codart & ";" & prezzo_str & vbCrLf & vbCrLf)

                        Catch ex As Exception
                            ' MsgBox(ex.Message)
                            ' MsgBox(cmdWmn.CommandText)
                        End Try
                    End If
                Next

                'Else
                '    For i = 1 To 30
                '        If flagTGVal(i) = 1 Then
                '            codice = codart & i.ToString("00")
                '            prezzo_str = Replace(rsteSList.Fields("PrezzoTaglia" & i.ToString()).Value.ToString(), ", ", ".")
                '            chiaveList = codice & "\" & rsteSList.Fields("CodiceNegozioSky").Value & "\" & rsteSList.Fields("CodValuta").Value & "\" & inizio_str & "\" & fine_str
                '            Try
                '                'metto codart al posto di codice (codart e poi codice)
                '                cmdWmn.CommandText = "INSERT INTO tiePrzAcquisto VALUES ('" & chiaveList & "', '" & codart & "', '" & codice & "', '" & rsteSList.Fields("CodiceNegozioSky").Value & "', '',1, '" & rsteSList.Fields("CodValuta").Value & "', " & prezzo_str & ",0,0,0,0,0,0, " & annullato.ToString() & ", 2, '" & str_dataUltAgg & "', 0 , '')"
                '                cmdWmn.CommandType = ADODB.CommandTypeEnum.adCmdText
                '                cmdWmn.Execute()
                '            Catch ex As Exception
                '                ' MsgBox(ex.Message)
                '                '  MsgBox(cmdWmn.CommandText)
                '            End Try
                '        End If
                '    Next
            End If
        End If
    End Sub

    Private Sub FormMain_Load(sender As Object, e As EventArgs) Handles MyBase.Shown


        Call connectToDB()

        Dim strArg() As String
        Dim inifilname As String
        Dim stazione As String
        strArg = Command().Split(" ")

        inifilname = strArg(0)
        stazione = strArg(1)

        'MsgBox(stazione)

        'MsgBox(inifilname)

        readIniFile(inifilname)



        nameFile = "LogSkyPrz" + stazione + ".log"

        'MsgBox(nameFile)

        Call LanciaEsportazioneSingola()


        Application.Exit()


    End Sub

End Class