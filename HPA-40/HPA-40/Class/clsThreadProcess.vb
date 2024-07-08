' HPA-40
' frmMain.vb
' ThreadProcessing
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

Public Class clsThreadProcess
#Region "ENUM"
#Region "ENUM-R/Wsign"
    Private Enum RWsign
        SGN_READ = 0      'Read
        SGN_WRITE         'Write
    End Enum
#End Region
#End Region

#Region "構造体"
#Region "構造体_ソケット通信設定値"
    Dim intSktPort As Integer = 502      'ポート番号
    Dim intSktIntval As Integer = 200    '送信間隔
    Dim intSktTimeout As Integer = 100   'タイムアウト

#End Region
#Region "構造体_Modbus書込みデータ"
    Private Structure sttWriteData
        Dim intWDataLoRaAddr As Integer 'LoRaアドレス
        Dim intWDataModAddr As Integer  'Modbusアドレス
        Dim intWDataYear As Integer     '年
        Dim intWDataMonth As Integer    '月
        Dim intWDataDay As Integer      '日
        Dim intWDataHour As Integer     '時
        Dim intWDataMin As Integer      '分
        Dim intWDataSec As Integer      '秒
    End Structure
#End Region
#End Region

#Region "変数"
    Private gobjClsTcpClient As New clsTcpClient   'clsTcpClientクラス
    Private gsttModbus As sttModbus
    Private gsttWriteData As sttWriteData
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

    Private strDataWriteForGW1 As String
    Private strDataWriteForGW2 As String
    Private strDataWriteForGW3 As String
    Private strDataWriteForGW4 As String
    Private strDataWriteForGW5 As String
    Private strDataWriteForGW6 As String
    Private strDataWriteForGW7 As String
    Private strDataWriteForGW8 As String
#End Region

#Region "関数_入力値をグローバル変数に入れる"
    '「IPアドレス」入力値チェック
    'Parameters:
    '  None
    'Returns:
    '  None
    Private Sub subPutSettingData(ByVal gwIndex As Integer, ByVal unitIndex As Integer)
        'テキストボックス_IPアドレス
        Dim gwIndexTemp As Integer = gwIndex
        Dim unitIndexTemp As Integer = unitIndex

        '転送ID
        gsttModbus.intModTransId = CInt("&H" & "0000")
        'プロトコルID
        gsttModbus.intModProtId = CInt("&H" & "0000")
        'ユニットID
        gsttModbus.intModUnitId = CInt("&H" & "00")
        'Read開始アドレス
        gsttModbus.intModReadAddr = 4109
        'Readレジスタ数
        gsttModbus.intModReadRegist = 72
        'テキストボックス_Write開始アドレス
        gsttModbus.intModWriteAddr = 1000
        'テキストボックス_Writeレジスタ数
        gsttModbus.intModWriteRegist = 8
        'GW1のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 0 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW1(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW1(unitIndexTemp)
        End If

        'GW2のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 1 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW2(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW2(unitIndexTemp)
        End If

        'GW3のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 2 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW3(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW3(unitIndexTemp)
        End If

        'GW4のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 3 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW4(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW4(unitIndexTemp)
        End If

        'GW5のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 4 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW5(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW5(unitIndexTemp)
        End If

        'GW6のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 5 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW6(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW6(unitIndexTemp)
        End If

        'GW7のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 6 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW7(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW7(unitIndexTemp)
        End If

        'GW8のLoRaアドレスおよびModbusアドレス
        If gwIndexTemp = 7 Then
            gsttWriteData.intWDataLoRaAddr = arrLoRaAddOfGW8(unitIndexTemp)
            gsttWriteData.intWDataModAddr = arrModAddOfGW8(unitIndexTemp)
        End If

        gsttModbus.intModFunc = &H17

        'Queueから時間データ
        Dim formatTime As String = "yyyy/MM/dd HH:mm:ss"
        Dim dateTimeConvert As DateTime = DateTime.ParseExact(dtTimeTaking, formatTime, CultureInfo.InvariantCulture)
        '年
        gsttWriteData.intWDataYear = dateTimeConvert.Year
        '月
        gsttWriteData.intWDataMonth = dateTimeConvert.Month
        '日
        gsttWriteData.intWDataDay = dateTimeConvert.Day
        '時
        gsttWriteData.intWDataHour = dateTimeConvert.Hour
        '分
        gsttWriteData.intWDataMin = dateTimeConvert.Minute
        '秒
        gsttWriteData.intWDataSec = dateTimeConvert.Second
    End Sub
#End Region

#Region "関数_送信処理"
    Private Function funcGWProcessing(ByVal gwIndex As Integer, ByVal unitIndex As Integer) As Boolean
        Dim blnRslt As Boolean = False
        Dim gwIndexTemp As Integer = gwIndex
        Dim unitIndexTemp As Integer = unitIndex

        Console.WriteLine("gwIndex: " & gwIndexTemp + 1)
        Console.WriteLine("unitIndex: " & unitIndexTemp + 1)

        '全てのデータを得る際の現在の送信回数がmaxを超えていれば初期化する
        If (gintComNowCnt >= gintComMaxCnt) Then
            gintComNowCnt = 0
        End If

        '全てのデータを得る際の現在の送信回数がmaxを超えていれば初期化する
        If (gintComNowCnt_w >= gintComMaxCnt) Then
            gintComNowCnt_w = 0
        End If

        'データ作成
        Dim intSndDataLen As Integer
        'レジスタ数が多きときに自動で分割する
        intSndDataLen = funcGetSndDatDev(gwIndexTemp, unitIndexTemp)
        If intSndDataLen = 0 Then
            Return blnRslt
            Exit Function
        End If

        '送信データをテキストボックスに書込む
        Dim strSend As String = ""
        For intCnt = 0 To (intSndDataLen - 1)
            strSend = strSend & "[" & gbytSndData(intCnt).ToString("X2") & "h] "
        Next

        'データを送信する
        blnRslt = gobjClsTcpClient.mTcpSend(gbytSndData, intSndDataLen)
        Thread.Sleep(500)
        '送信処理が正常に行えなかったときは、終了処理を行う
        If blnRslt = False Then
            Return blnRslt
            Exit Function
        End If

        '全てのデータを得る際の現在の送信回数をカウントアップ
        gintComNowCnt += 1
        '受信確認用タイマスタート
        gintCntTmr1Tick = 0
        '受信処理
        blnRslt = gobjClsTcpClient.mTcpRcv
        Thread.Sleep(200)

        Dim intRcvStat As Integer
        'Tickが起こった回数をカウントアップ
        gintCntTmr1Tick += 1
        '受信処理
        blnRslt = gobjClsTcpClient.mTcpRcv
        Thread.Sleep(200)
        '受信状態フラグ確認
        intRcvStat = gobjClsTcpClient.pRcvStat
        If intRcvStat = RcvStat.RCV_END Then
            Dim dtmNow As DateTime
            dtmNow = DateTime.Now
            '受信完了
            'Tickが起こった回数を初期化する
            gintCntTmr1Tick = 0
            '受信データを取得する
            Dim bytRcvData() As Byte
            bytRcvData = gobjClsTcpClient.pRcvData
            '---受信データを[RawData]タブのテキストボックスに書込む---
            Dim strRcv As String = ""
            Dim strRcv1 As String = ""
            For intCnt = 0 To (bytRcvData.Length - 1)
                strRcv = strRcv & bytRcvData(intCnt).ToString("X2")
            Next

            For intCnt = 0 To (bytRcvData.Length - 1)
                strRcv1 = strRcv1 & bytRcvData(intCnt).ToString("X2")
            Next
            Dim data As String = strRcv1

            '最後に改行を追加する
            Dim strDataFilter As String = (strRcv).Substring(18)

            'データラインを作成
            Dim strDataLine As String
            If unitIndexTemp = 0 Then
                strDataLine = dtmNow.ToString("yyyy/MM/dd HH:mm:ss.fff") & "," & dtTimeTaking & funcCreateDataLine(strDataFilter)
            Else
                strDataLine = funcCreateDataLine(strDataFilter)
            End If

            'Queueから時間データ
            Dim formatTime As String = "yyyy/MM/dd HH:mm:ss"
            Dim dateTimeConvert As DateTime = DateTime.ParseExact(dtTimeTaking, formatTime, CultureInfo.InvariantCulture)

            'CSVファイル名を作成
            Dim strCSVFilePath As String
            Dim strCTCFilePath As String
            Dim strTimePart As String

            strTimePart = dateTimeConvert.Year.ToString("0000") & dateTimeConvert.Month.ToString("00") & _
                            dateTimeConvert.Day.ToString("00") & dateTimeConvert.Hour.ToString("00") & _
                            dateTimeConvert.Minute.ToString("00") & dateTimeConvert.Second.ToString("00")

            If gwIndexTemp = 0 Then
                strDataWriteForGW1 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw1.UnitNumber - 1 Then
                    'ファイル名を設定
                    strCSVFilePath = strbasePathGateway & gw1.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw1.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW1)
                    'データラインをリセット
                    strDataWriteForGW1 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 1 Then
                strDataWriteForGW2 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw2.UnitNumber - 1 Then
                    'ファイル名を設定
                    strCSVFilePath = strbasePathGateway & gw2.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw2.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW2)
                    'データラインをリセット
                    strDataWriteForGW2 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 2 Then
                strDataWriteForGW3 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw3.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw3.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw3.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW3)
                    'データラインをリセット
                    strDataWriteForGW3 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 3 Then
                strDataWriteForGW4 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw4.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw4.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw4.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW4)
                    'データラインをリセット
                    strDataWriteForGW4 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 4 Then
                strDataWriteForGW5 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw5.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw5.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw5.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW5)
                    'データラインをリセット
                    strDataWriteForGW5 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 5 Then
                strDataWriteForGW6 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw6.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw6.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw6.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW6)
                    'データラインをリセット
                    strDataWriteForGW6 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            ElseIf gwIndexTemp = 6 Then
                strDataWriteForGW7 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw7.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw7.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw7.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW7)
                    'データラインをリセット
                    strDataWriteForGW7 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If

            Else
                strDataWriteForGW8 += strDataLine
                '最後のユニットになったかどうか確認
                If unitIndexTemp = gw8.UnitNumber - 1 Then
                    strCSVFilePath = strbasePathGateway & gw8.GWName & "_" & strTimePart & ".CSV"
                    strCTCFilePath = strbasePathGateway & gw8.GWName & "_" & strTimePart & ".CTC"
                    'CSVファイルを作成しデータを書き込む
                    subCreateAndWriteCSV(strCSVFilePath, strDataWriteForGW8)
                    'データラインをリセット
                    strDataWriteForGW8 = ""
                    'CTCファイルを作成
                    subCreateFileCTC(strCTCFilePath)
                End If
            End If

            '受信データチェック
            blnRslt = funcChkRcvData(bytRcvData)
            If blnRslt = False Then
                '受信エラー処理
                SubProcRcvErr()
                Dim intCnt As String = bytRcvData.Length - 1
                MES_RCV_ERR_2 = "エラーコード" + bytRcvData(intCnt).ToString("X2") & "H" + " :受信できません"
                Exit Function
            End If

            'データを分けて送受信する場合は、送信間隔に関わらず送信する
            If gintComNowCnt < gintComMaxCnt Then
                'データ送信処理
                blnRslt = funcGWProcessing(gwIndexTemp, unitIndexTemp)
                '送信処理が正常に行えなかったときは、メッセージを表示する
                If blnRslt = False Then
                    '「TCP接続が出来ません」エラーを記録する
                End If
                '送信間隔管理用タイマスタートする
                If ginitLoop = True Then
                    '送信間隔管理用タイマスタート
                    gintCntTmr2Tick = 0
                End If
            Else
                'chkLoopにチェックが入っていなければ、停止処理を行う
                If ginitLoop = False Then
                    'サーバ側との接続を切断する
                    gobjClsTcpClient.mTcpClose()
                End If
            End If
        ElseIf intRcvStat = RcvStat.RCV_MID Then
            '受信中
            If gintCntTmr1Tick >= (intSktTimeout \ intSktIntval) Then
                'タイムアウトした
                'Tickが起こった回数を初期化する
                gintCntTmr1Tick = 0
                '受信エラー処理
                SubProcRcvErr()
            Else
                'タイムアウトがまだ→タイマスタート
                funcGWProcessing(gwIndexTemp, unitIndexTemp)
            End If
        ElseIf intRcvStat = RcvStat.RCV_ERR Then
            '受信エラー
            'Tickが起こった回数を初期化する
            gintCntTmr1Tick = 0
            '受信エラー処理
            SubProcRcvErr()
            Exit Function
        End If
        blnRslt = True
        Return blnRslt
    End Function
#End Region

#Region "チャンネル数をチェックし、データラインを作成"
    Private Function funcCreateDataLine(ByVal strDataFilter As String)
        Dim strDataFilterTemp As String = strDataFilter
        Dim strDataLine As String
        Dim strRevertData1 As String
        Dim decChannel1 As Single
        Dim strRevertData2 As String
        Dim decChannel2 As Single
        Dim strRevertData3 As String
        Dim decChannel3 As Single
        Dim strRevertData4 As String
        Dim decChannel4 As Single
        Dim strRevertData5 As String
        Dim decChannel5 As Single
        Dim strRevertData6 As String
        Dim decChannel6 As Single
        Dim strRevertData7 As String
        Dim decChannel7 As Single
        Dim strRevertData8 As String
        Dim decChannel8 As Single
        Dim strdecChannel1 As String
        Dim strdecChannel2 As String
        Dim strdecChannel3 As String
        Dim strdecChannel4 As String
        Dim strdecChannel5 As String
        Dim strdecChannel6 As String
        Dim strdecChannel7 As String
        Dim strdecChannel8 As String

        Dim strStatusChannel1 As String = strDataFilterTemp.Substring(0, 4)
        Dim strStatusChannel2 As String = strDataFilterTemp.Substring(36, 4)
        Dim strStatusChannel3 As String = strDataFilterTemp.Substring(72, 4)
        Dim strStatusChannel4 As String = strDataFilterTemp.Substring(108, 4)
        Dim strStatusChannel5 As String = strDataFilterTemp.Substring(144, 4)
        Dim strStatusChannel6 As String = strDataFilterTemp.Substring(180, 4)
        Dim strStatusChannel7 As String = strDataFilterTemp.Substring(216, 4)
        Dim strStatusChannel8 As String = strDataFilterTemp.Substring(252, 4)

        'Channel1
        'アナログ入力が不使用かどうか確認
        If strStatusChannel1 <> "0000" Then
            '16進数
            strRevertData1 = strDataFilterTemp.Substring(20, 16)
            '10進数に変更
            decChannel1 = Math.Round(Convert.ToInt32(strRevertData1, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel1 = decChannel1.ToString("F2")
        Else
            strdecChannel1 = ""
        End If

        'Channel2
        'アナログ入力が不使用かどうか確認
        If strStatusChannel2 <> "0000" Then
            '16進数
            strRevertData2 = strDataFilterTemp.Substring(56, 16)
            '10進数に変更
            decChannel2 = Math.Round(Convert.ToInt32(strRevertData2, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel2 = decChannel2.ToString("F2")
        Else
            strdecChannel2 = ""
        End If

        'Channel3
        'アナログ入力が不使用かどうか確認
        If strStatusChannel3 <> "0000" Then
            '16進数
            strRevertData3 = strDataFilterTemp.Substring(92, 16)
            '10進数に変更
            decChannel3 = Math.Round(Convert.ToInt32(strRevertData3, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel3 = decChannel3.ToString("F2")
        Else
            strdecChannel3 = ""
        End If

        'Channel4
        'アナログ入力が不使用かどうか確認
        If strStatusChannel4 <> "0000" Then
            '16進数
            strRevertData4 = strDataFilterTemp.Substring(128, 16)
            '10進数に変更
            decChannel4 = Math.Round(Convert.ToInt32(strRevertData4, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel4 = decChannel4.ToString("F2")
        Else
            strdecChannel4 = ""
        End If

        'Channel5
        'アナログ入力が不使用かどうか確認
        If strStatusChannel5 <> "0000" Then
            '16進数
            strRevertData5 = strDataFilterTemp.Substring(164, 16)
            '10進数に変更
            decChannel5 = Math.Round(Convert.ToInt32(strRevertData5, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel5 = decChannel5.ToString("F2")
        Else
            strdecChannel5 = ""
        End If

        'Channel6
        'アナログ入力が不使用かどうか確認
        If strStatusChannel6 <> "0000" Then
            '16進数
            strRevertData6 = strDataFilterTemp.Substring(200, 16)
            '10進数に変更
            decChannel6 = Math.Round(Convert.ToInt32(strRevertData6, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel6 = decChannel6.ToString("F2")
        Else
            strdecChannel6 = ""
        End If

        'Channel7
        'アナログ入力が不使用かどうか確認
        If strStatusChannel7 <> "0000" Then
            '16進数
            strRevertData7 = strDataFilterTemp.Substring(236, 16)
            '10進数に変更
            decChannel7 = Math.Round(Convert.ToInt32(strRevertData7, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel7 = decChannel7.ToString("F2")
        Else
            strdecChannel7 = ""
        End If

        'Channel8
        'アナログ入力が不使用かどうか確認
        If strStatusChannel8 <> "0000" Then
            '16進数
            strRevertData8 = strDataFilterTemp.Substring(272, 16)
            '10進数に変更
            decChannel8 = Math.Round(Convert.ToInt32(strRevertData8, 16) * 0.0001, 2)
            '小数点(2桁)を取る
            strdecChannel8 = decChannel8.ToString("F2")
        Else
            strdecChannel8 = ""
        End If

        '時間の時刻と電流瞬時値のデータを組み合わせる
        strDataLine = "," & strdecChannel1 & "," & strdecChannel2 &
                    "," & strdecChannel3 & "," & strdecChannel4 &
                    "," & strdecChannel5 & "," & strdecChannel6 &
                    "," & strdecChannel7 & "," & strdecChannel8

        Return strDataLine
    End Function
#End Region

#Region "関数_CSVファイルを作成しデータを書き込む"
    Private Sub subCreateAndWriteCSV(ByVal strFilePathCSV As String, ByVal strDataLine As String)
        Dim strFilePathCSVTemp As String = strFilePathCSV
        Dim strDataLineTemp As String = strDataLine
        'CSVファイルを作成
        Using fsCreateCSV As FileStream = File.Create(strFilePathCSVTemp)
        End Using
        'CSVファイルにデータを書き込む
        If File.Exists(strFilePathCSVTemp) Then
            Dim writerDataCSV As StreamWriter = Nothing
            'Encoding: SHIFT-JIS
            writerDataCSV = New StreamWriter(strFilePathCSVTemp, True, System.Text.Encoding.GetEncoding("Shift-JIS"))
            writerDataCSV.WriteLine(strDataLineTemp)
            If writerDataCSV IsNot Nothing Then
                'リソースを解放する
                writerDataCSV.Dispose()
            End If
        End If
    End Sub
#End Region

#Region "関数_CTCファイルを作成"
    Private Function subCreateFileCTC(strFilePathCTC As String) As Boolean
        Dim strFilePathCTCTemp As String = strFilePathCTC
        'CTCファイルを作成
        Using fsCreateCTC As FileStream = File.Create(strFilePathCTCTemp)
            Return True
        End Using
        Return False
    End Function
#End Region

#Region "関数_送信データ生成_レジスタ数が126以上で指定されていたら通信を2回に分ける"
    Private Function funcGetSndDatDev(ByVal gwIndex As Integer, ByVal unitIndex As Integer) As Integer
        Dim intDataLen As Integer
        Dim intBuf As Integer
        Dim gwIndexTemp As Integer = gwIndex
        Dim unitIndexTemp As Integer = unitIndex

        '共通処理_転送ID
        gbytSndData(0) = (gsttModbus.intModTransId And &HFF00) >> 8
        gbytSndData(1) = gsttModbus.intModTransId And &HFF
        '共通処理_プロトコルID
        gbytSndData(2) = (gsttModbus.intModProtId And &HFF00) >> 8
        gbytSndData(3) = gsttModbus.intModProtId And &HFF
        Dim intLenBuf As Integer
        intLenBuf = 0

        '共通処理_ユニットID
        gbytSndData(6 + intLenBuf) = gsttModbus.intModUnitId
        intLenBuf += 1
        'ファンクションコード
        gbytSndData(6 + intLenBuf) = gsttModbus.intModFunc
        intLenBuf += 1
        '読み込み開始アドレス
        intBuf = gsttModbus.intModReadAddr + (gintComNowCnt * REGI_ONE_MAX)
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '読み込みレジスタ数
        If gsttModbus.intModReadRegist > ((gintComNowCnt + 1) * REGI_ONE_MAX) Then
            intBuf = REGI_ONE_MAX
        Else
            intBuf = gsttModbus.intModReadRegist - (gintComNowCnt * REGI_ONE_MAX)
        End If
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込む開始アドレス
        intBuf = gsttModbus.intModWriteAddr
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1

        '指定した時刻等を書き込む
        'レジスタ数
        intBuf = TIME_WRITE_DATA_NUM
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        'データバイト数
        intBuf = intBuf * 2
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ
        intLenBuf = funcGetSndTimeData(intLenBuf, gwIndexTemp, unitIndexTemp)

        '転送バイト数
        gbytSndData(4) = (intLenBuf And &HFF00) >> 8
        gbytSndData(5) = intLenBuf And &HFF
        intDataLen = 6 + intLenBuf

        Return intDataLen
    End Function
#End Region

#Region "関数_送信データ時間データ生成"
    '送信データ時間データ生成
    'Parameters:
    '  ByVal intLenBuf As Integer：オフセット
    'Returns:
    '  Integer：送信電文長
    Private Function funcGetSndTimeData(ByVal intLenBuf As Integer, ByVal gwIndex As Integer, ByVal unitIndex As Integer) As Integer
        Dim intBuf As Integer
        Dim gwIndexTemp As Integer = gwIndex
        Dim unitIndexTemp As Integer = unitIndex

        '書き込むデータ_LoRaアドレス
        intBuf = gsttWriteData.intWDataLoRaAddr
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_Modbusアドレス
        intBuf = gsttWriteData.intWDataModAddr
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_年
        intBuf = gsttWriteData.intWDataYear
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_月
        intBuf = gsttWriteData.intWDataMonth
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_日
        intBuf = gsttWriteData.intWDataDay
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_時
        intBuf = gsttWriteData.intWDataHour
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_分
        intBuf = gsttWriteData.intWDataMin
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        '書き込むデータ_秒
        intBuf = gsttWriteData.intWDataSec
        gbytSndData(6 + intLenBuf) = (intBuf And &HFF00) >> 8
        intLenBuf += 1
        gbytSndData(6 + intLenBuf) = intBuf And &HFF
        intLenBuf += 1
        Return intLenBuf
    End Function
#End Region

#Region "関数_受信データチェック"
    '受信データチェック
    'Parameters:
    '  ByVal bytRcvData() As Byte：受信データ
    'Returns:
    '  Boolean：True=OK, False=NG
    Private Function funcChkRcvData(ByVal bytRcvData() As Byte) As Boolean
        Dim blnRslt As Boolean = False
        Dim intBuf As Integer
        Dim intRegNum As Integer
        '転送IDチェック
        intBuf = (CInt(bytRcvData(0)) << 8)
        intBuf = intBuf + bytRcvData(1)
        If intBuf <> gsttModbus.intModTransId Then
            Return blnRslt
            Exit Function
        End If
        'プロトコルIDチェック
        intBuf = (CInt(bytRcvData(2)) << 8)
        intBuf = intBuf + bytRcvData(3)
        If intBuf <> gsttModbus.intModProtId Then
            Return blnRslt
            Exit Function
        End If
        '転送バイト数チェック
        intBuf = (CInt(bytRcvData(4)) << 8)
        intBuf = intBuf + bytRcvData(5)
        If intBuf <> (bytRcvData.Length - 6) Then
            Return blnRslt
            Exit Function
        End If
        'ユニットIDチェック
        intBuf = bytRcvData(6)
        If intBuf <> gsttModbus.intModUnitId Then
            Return blnRslt
            Exit Function
        End If
        'ファンクションコードチェック
        intBuf = bytRcvData(7)
        If intBuf <> gsttModbus.intModFunc Then
            Return blnRslt
            Exit Function
        End If
        '以下はファンクションによって異なる
        '読み出しコマンドの場合は受信データの「バイト数」の項目が、要求したレジスタ数の2倍であることを確認する
        intBuf = bytRcvData(8)
        If gsttModbus.intModReadRegist > (gintComNowCnt * REGI_ONE_MAX) Then
            intRegNum = REGI_ONE_MAX
        Else
            intRegNum = gsttModbus.intModReadRegist - ((gintComNowCnt - 1) * REGI_ONE_MAX)
        End If
        If intBuf <> (intRegNum * 2) Then
            Return blnRslt
            Exit Function
        End If

        blnRslt = True
        Return blnRslt
    End Function
#End Region

#Region "関数_受信エラー処理"
    '受信エラー処理
    'Parameters:
    '  None
    'Returns:
    '  None
    Private Sub SubProcRcvErr()
        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt = 0
        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt_w = 0
        'chkLoopにチェックが入っていなければ、停止処理を行う
        If ginitLoop = False Then
            'サーバ側との接続を切断する
            gobjClsTcpClient.mTcpClose()
        End If
    End Sub
#End Region

#Region "関数_Threadの処理"
    Public Sub subProcessingThread(ByVal indexGW As Integer, ByVal indexUnit As Integer)
        Dim indexGWTemp As Integer = indexGW
        Dim indexUnitTemp As Integer = indexUnit
        Dim blnRslt As Boolean
        '入力値をグローバル変数に書込む
        subPutSettingData(indexGWTemp, indexUnitTemp)

        'レジスタ数が多きときに自動で分割して送信する
        gintComMaxCnt = (gsttModbus.intModReadRegist \ (REGI_ONE_MAX + 1)) + 1

        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt = 0
        '全てのデータを得る際の現在の送信回数を初期化する
        gintComNowCnt_w = 0

        'IPAddressの情報を取得
        Dim strIPAddress As String
        If indexGWTemp = 0 Then
            strIPAddress = gw1.IPAddress
        ElseIf indexGWTemp = 1 Then
            strIPAddress = gw2.IPAddress
        ElseIf indexGWTemp = 2 Then
            strIPAddress = gw3.IPAddress
        ElseIf indexGWTemp = 3 Then
            strIPAddress = gw4.IPAddress
        ElseIf indexGWTemp = 4 Then
            strIPAddress = gw5.IPAddress
        ElseIf indexGWTemp = 5 Then
            strIPAddress = gw6.IPAddress
        ElseIf indexGWTemp = 6 Then
            strIPAddress = gw7.IPAddress
        Else
            strIPAddress = gw8.IPAddress
        End If

        'TCP接続を行う
        blnRslt = gobjClsTcpClient.mTcpConnect(strIPAddress, intSktPort, intSktTimeout)

        'TCP接続ができなかったときは、終了処理を行う
        If blnRslt = False Then
            Console.WriteLine("TCP接続できません")
            Exit Sub
        End If

        'データ送信処理、受信、CSV書き込み
        blnRslt = funcGWProcessing(indexGWTemp, indexUnitTemp)

        If blnRslt = False Then
            Console.WriteLine("処理はエラーが発生しています")
        End If
    End Sub
#End Region

End Class
