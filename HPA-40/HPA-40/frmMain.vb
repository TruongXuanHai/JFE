' HPA-40
' frmMain.vb
' ユーザーインターフェース
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports HPA_40.clsAllVariable
Imports HPA_40.clsIniFunc
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Threading.Tasks
Imports System.Threading
Imports System.Globalization
Imports System.Timers
Imports MS.Internal.Xml
Imports HPA_40.clsGetParameter

Public Class frmMain
#Region "変数"
    Private gobjClsTcpClient As New clsTcpClient   'clsTcpClientクラス
    Private gsttModbus As sttModbus
    Private gintCntTmr1Tick As Integer             'tmrTimer1のTickが起こった回数
    Private gintCntTmr2Tick As Integer             'tmrTimer2のTickが起こった回数
    Private gbytSndData(32768) As Byte             '送信データ
    Private gintComMaxCnt As Integer               '全てのデータを得るのに必要な送信回数[（ユーザーの指定レジスタ数 \ REGI_ONE_MAX）+ 1]
    Private gintComNowCnt As Integer = 0           '全てのデータを得る際の現在の送信回数
    Private gintComNowCnt_w As Integer = 0         '全てのデータを得る際の現在の送信回数(Write用)
    Private ginitLoop As Boolean = False           'Loop許可変数 (True→Loop許可, False→Loop許可なし)
    Private gLongReg As Boolean = False            'レジスタ数を分ける許可変数 (False→レジスタ数が多きときに自動で分割して送信する, True→レジスタ数が多きときに自動で分割して送信する)
    Private gTimeWrite As Boolean = True           'Time追加許可変数
    'コントロール管理
    Private gtxtReadData(REGI_ALL_MAX - 1) As TextBox
    Private gtxtWriteData(REGI_ALL_MAX - 1) As TextBox

    'クラスを宣言
    Private taskGateway As New clsTaskGateway()
    Private taskServer As New clsTaskServer()
    Private getParameter As New clsGetParameter()
#End Region

#Region "イベント_起動"
    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'ディレクトリを定義
        Dim objFileInfo As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        strbasePathGateway = objFileInfo.DirectoryName & "\Gateway\"                               'Gatewayフォルダ
        strbasePathServer = objFileInfo.DirectoryName & "\Server\"                                 'Serverフォルダ
        strbasePathFailed = objFileInfo.DirectoryName & "\Failed\"                                 'Failedフォルダ
        strbasePathErrLog = objFileInfo.DirectoryName & "\ErrLog\"                                 'ErrLogフォルダ
        strxmlFilePathGateway = objFileInfo.DirectoryName & "\" & UNITSETTING_XML_NAME & ".xml"    'GW通信設定XMLパス
        strxmlFilePathServer = objFileInfo.DirectoryName & "\" & FTPSETTING_XML_NAME & ".xml"      'Server通信設定XMLパス
        strswitchONPath = objFileInfo.DirectoryName & "\Icon\SWITCH-ON.png"                        'ONスウィッチイメージパス
        strswitchOFFPath = objFileInfo.DirectoryName & "\Icon\SWITCH-OFF.png"                      'OFFスウィッチイメージパス

        'XMLファイルを読み出し変数を格納
        getParameter.subLoadXMLSetting(strxmlFilePathGateway, strxmlFilePathServer)

        'Threadの初期チェック用フラグをFalseにする
        gblthreadsInitialized = False



        'Queueを初期
        timeQueue = New Queue(Of String)()
        'Timerを初期, Interval: 1分 (60000ms)
        timerForQueue = New System.Timers.Timer(60000)
        AddHandler timerForQueue.Elapsed, AddressOf subOnTimedEvent
        timerForQueue.AutoReset = True
        timerForQueue.Enabled = True
        subOnTimedEvent(Nothing, Nothing)

        'フォーム初期化
        subFormInit(gwSettingInfo, serverSettingInfo)

        'Task1を作成
        taskGateway.subInitializeTaskGateway()
        'Task2を作成
        taskServer.subInitializeTaskServer()
    End Sub
#End Region

#Region "関数_本アプリの起動ディレクトリパスを取得する"
    '本アプリの起動ディレクトリパスを取得する
    Private Function funcGetAppPath() As String
        Dim objFileInfo As New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Return objFileInfo.DirectoryName
    End Function
#End Region

#Region "関数_フォーム初期化"
    'フォーム初期化
    Public Sub subFormInit(ByVal gwSettingTemp As Integer, ByVal serverSettingTemp As Integer)
        'フォーム名設定
        Me.Text = TITLE
        'Loop用のチェックボックス
        ginitLoop = False
        '126バイト以上のデータ通信許可
        gLongReg = False   '許可しない
        'ファンクションコード固定
        gsttModbus.intModFunc = &H17
        '書込みデータ固定
        gTimeWrite = True

        'GW通信状態更新
        If gwSettingTemp = 0 Then
            pbGWStatus.Image = Image.FromFile(strswitchOFFPath)
            lblGWStatus.Text = "OFF"

            'GW通信処理用終了フラフ
            gblStopFlagGateway = False
            'GW通信処理用中断フラフ
            gblStopTempFlagGateway = True

        ElseIf gwSettingTemp = 1 Then
            pbGWStatus.Image = Image.FromFile(strswitchONPath)
            lblGWStatus.Text = "ON"

            'GW通信処理用終了フラフ
            gblStopFlagGateway = False
            'GW通信処理用中断フラフ
            gblStopTempFlagGateway = False
        End If

        'サーバ通信状態更新
        If serverSettingTemp = 0 Then
            pbServerStatus.Image = Image.FromFile(strswitchOFFPath)
            lblServerStatus.Text = "OFF"
            'サーバ通信処理用終了フラフ
            gblStopFlagServer = False
            'サーバ通信処理用中断フラフ
            gblStopTempFlagServer = True
        ElseIf serverSettingTemp = 1 Then
            pbServerStatus.Image = Image.FromFile(strswitchONPath)
            lblServerStatus.Text = "ON"
            'サーバ通信処理用終了フラフ
            gblStopFlagServer = False
            'サーバ通信処理用中断フラフ
            gblStopTempFlagServer = False
        End If

        '1秒毎クロックの時間を更新
        tmrTimeClock.Interval = 1000 '1秒
        AddHandler tmrTimeClock.Tick, AddressOf tmrTimeClock_Tick
        tmrTimeClock.Start()

        '開始時間及び終了時間のフォーマットを設定
        With dtpStartTime
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "yyyy MM dd HH"
        End With
        '終了時間及び終了時間のフォーマットを設定
        With dtpEndTime
            .Format = DateTimePickerFormat.Custom
            .CustomFormat = "yyyy MM dd HH"
        End With

        'Gatewayフォルダを作成
        If Not Directory.Exists(strbasePathGateway) Then
            Directory.CreateDirectory(strbasePathGateway)
        End If

        'Serverフォルダを作成
        If Not Directory.Exists(strbasePathServer) Then
            Directory.CreateDirectory(strbasePathServer)
        End If

        'Failedフォルダを作成
        If Not Directory.Exists(strbasePathFailed) Then
            Directory.CreateDirectory(strbasePathFailed)
        End If

        'ErrLogフォルダを作成
        If Not Directory.Exists(strbasePathErrLog) Then
            Directory.CreateDirectory(strbasePathErrLog)
        End If
    End Sub
#End Region

#Region "クロックタイム表示"
    '定期時間（10分毎に）時間データをQueueに入れる
    Public Sub subOnTimedEvent(source As Object, e As ElapsedEventArgs)
        Dim dtTimeNow As DateTime = DateTime.Now
        '時間が0分、10分、20分、30分、40分、50分かを確認
        If dtTimeNow.Minute Mod 10 = 0 Then
            '現在時刻より10分の遅延を設定
            Dim dtTimeGW As DateTime = dtTimeNow.AddMinutes(-10)
            '"YYYY/MM/DD hh:mm:00"フォーマットを設定
            Dim dtTimeGWMod As String = dtTimeGW.ToString("yyyy/MM/dd HH:mm:00")
            'Queueにデータ時間を追加
            timeQueue.Enqueue(dtTimeGWMod)
        End If
    End Sub
#End Region

#Region "クロックタイム表示"
    '現在時刻とQueueのカウンターを表示する
    Private Sub tmrTimeClock_Tick(sender As Object, e As EventArgs)
        lblTimeClock.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")
        lblPendingNumber.Text = timeQueue.Count
    End Sub
#End Region

#Region "イベント_GW通信を押下"
    Private Sub pbGWStatus_Click(sender As Object, e As EventArgs) Handles pbGWStatus.Click
        subUpdateGateway()
    End Sub
#End Region

#Region "関数_GW通信を押下した後コントロール処理"
    'PictureBoxとLabel(ON/OFF)を切り替える
    'GatewayTaskにて指示を出す
    Private Sub subUpdateGateway()
        If lblGWStatus.Text = "OFF" Then
            lblGWStatus.Text = "ON"
            pbGWStatus.Image = Image.FromFile(strswitchONPath)

            '中断フラグをFalseに設定
            gblStopTempFlagGateway = False

        ElseIf lblGWStatus.Text = "ON" Then
            lblGWStatus.Text = "OFF"
            pbGWStatus.Image = Image.FromFile(strswitchOFFPath)

            '中断フラグをTrueに設定
            gblStopTempFlagGateway = True
        End If
    End Sub
#End Region

#Region "イベント_サーバ通信を押下"
    Private Sub pbServerStatus_Click(sender As Object, e As EventArgs) Handles pbServerStatus.Click
        subUpdateServer()
    End Sub
#End Region

#Region "関数_サーバ通信を押下した後コントロール処理"
    'PictureBoxとLabel(ON/OFF)を切り替える
    'ServerTaskにて指示を出す
    Private Sub subUpdateServer()
        If lblServerStatus.Text = "OFF" Then
            lblServerStatus.Text = "ON"
            pbServerStatus.Image = Image.FromFile(strswitchONPath)

            '中断フラグをFalseに設定
            gblStopTempFlagServer = False
        ElseIf lblServerStatus.Text = "ON" Then
            lblServerStatus.Text = "OFF"
            pbServerStatus.Image = Image.FromFile(strswitchOFFPath)

            '中断フラグをTrueに設定
            gblStopTempFlagServer = True
        End If
    End Sub
#End Region

#Region "イベント_データ取得ボタン押下"
    Private Sub btnDataCollect_Click(sender As Object, e As EventArgs) Handles btnDataCollect.Click
        Dim startTime As DateTime = dtpStartTime.Value
        Dim endTime As DateTime = dtpEndTime.Value
        Dim startTimeMod = New DateTime(startTime.Year, startTime.Month, startTime.Day, startTime.Hour, 0, 0)
        Dim startTimeCompare = New DateTime(startTime.Year, startTime.Month, startTime.Day, 0, 0, 0)
        Dim endTimeMod = New DateTime(endTime.Year, endTime.Month, endTime.Day, endTime.Hour, 0, 0)
        Dim endTimeCompare = New DateTime(endTime.Year, endTime.Month, endTime.Day, 0, 0, 0)
        Dim timeDistance = endTimeCompare.Subtract(startTimeCompare)

        '時間エラーを確認
        If startTimeMod > DateTime.Now Then
            MessageBox.Show("開始時間が現在時刻より未来に設定されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf endTimeMod > DateTime.Now Then
            MessageBox.Show("終了時間が現在時刻より未来に設定されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf startTimeMod > endTimeMod Then
            MessageBox.Show("開始時間が終了時間より未来に設定されています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        ElseIf Math.Abs(timeDistance.TotalDays) > 31 Then
            MessageBox.Show("期間が１ヶ月を超えています。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            Dim currentTime As DateTime = startTimeMod
            Dim addTime As String
            '開始時間→終了時間の間：10分の間隔がある時間データをQueueに追加する
            While currentTime <= endTimeMod
                addTime = currentTime.ToString("yyyy/MM/dd HH:mm:00")
                timeQueue.Enqueue(addTime)
                currentTime = currentTime.AddMinutes(10)
            End While
        End If
    End Sub
#End Region

#Region "イベント_終了ボタン押下"
    Private Sub btnEnd_Click(sender As Object, e As EventArgs) Handles btnEnd.Click
        'Gateway用終了フラグを有効にする
        gblStopFlagGateway = True
        'サーバ用終了フラグを有効にする
        gblStopFlagServer = True
        '判定時間を超える場合Threadを強制敵に終了
        taskGateway.subStopThreads()
        '判定時間を超える場合TaskとThreadを強制敵に終了
        subStopAllTasks()
        'XMLファイルにGW通信とサーバ通信の状態を書き込む
        subUpdateGatewayAndServer()
        'ソフトをクローズ
        Me.Close()
    End Sub
#End Region

#Region "関数_Taskを強制的に終了する"
    Public Sub subStopAllTasks()
        If gtaskGatewayProcessing IsNot Nothing Then
            'Task1は20秒以内に終了できない場合、エラー扱いにする
            'Task.Wait(Int32)→スレッドが終了した場合は戻り値：True
            'millisecondsTimeoutパラメーターで指定した時間が経過してもTaskが終了していない場合は戻り値：False
            If Not gtaskGatewayProcessing.Wait(20000) Then
                Console.WriteLine("Task1が20秒以内に終了できません。強制的に終了します。")
                Me.Close()
            End If
        End If

        If gtaskServerProcessing IsNot Nothing Then
            'Task2は20秒以内に終了できない場合、エラー扱いにする
            'Task.Wait(Int32)→割り当てられた時間内に完了した場合戻り値：True
            'millisecondsTimeoutパラメーターで指定した時間が経過してもTaskが終了していない場合は戻り値：False
            If Not gtaskServerProcessing.Wait(20000) Then
                Console.WriteLine("Task2が20秒以内に終了できません。強制的に終了します。")
                Me.Close()
            End If
        End If
    End Sub
#End Region

#Region "関数_コントロール処理"
    'ソフト終了の時、XMLファイルにGW通信とサーバ通信の状態を書き込む
    Private Sub subUpdateGatewayAndServer()
        If lblGWStatus.Text = "OFF" Then
            getParameter.settings.GWSetting = 0
        Else
            getParameter.settings.GWSetting = 1
        End If
        If lblServerStatus.Text = "OFF" Then
            getParameter.settings.ServerSetting = 0
        Else
            getParameter.settings.ServerSetting = 1
        End If
        'XMLファイルにGW通信とサーバ通信の状態を書き込む
        getParameter.subWriteXMLSetting(strxmlFilePathGateway)
    End Sub
#End Region

#Region "イベント_フォームのクローズ"
    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        'Gateway用終了フラグを有効にする
        gblStopFlagGateway = True
        'サーバ用終了フラグを有効にする
        gblStopFlagServer = True
        '判定時間を超える場合Threadを強制敵に終了
        taskGateway.subStopThreads()
        '判定時間を超える場合Taskを強制敵に終了
        subStopAllTasks()
        'XMLファイルにGW通信とサーバ通信の状態を書き込む
        subUpdateGatewayAndServer()
    End Sub
#End Region
End Class

