' HPA-40
' mdGlobalVariable.vb
' XMLファイルから設定値を取得
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading.Tasks
Module mdGlobalVariable
#Region "ディレクトリ"
    Public strbasePathGateway As String         'Gatewayフォルダ
    Public strbasePathServer As String          'Serverフォルダ
    Public strbasePathFailed As String          'Failedフォルダ
    Public strbasePathErrLog As String          'ErrLogフォルダ
    Public strxmlFilePathGateway As String      'GW通信設定XMLパス
    Public strxmlFilePathServer As String       'Server通信設定XMLパス
    Public strswitchONPath As String            'ONスウィッチイメージパス
    Public strswitchOFFPath As String           'OFFスウィッチイメージパス
#End Region

#Region "固定変数"
    Public APP_NAME_1 As String = "HPA-40"
    Public UNITSETTING_XML_NAME As String = "Unitsetting1"
    Public FTPSETTING_XML_NAME As String = "Serversetting"
    Public MAINLOG_TXT_NAME As String = "MainErr"
    Public APP_NAME_2 As String = "上位サーバ転送"
    Public VER_NUM As String = "Ver.1.00"
    Public TITLE As String = APP_NAME_2 & " (" & APP_NAME_1 & ")" & " " & VER_NUM
    Public MES_RCV_OK_1 As String = "受信OK"
    Public MES_RCV_ERR_1 As String = "受信エラー：受信できません"
    Public MES_RCV_ERR_2 As String = ""
    Public MES_SND_ERR_1 As String = "TCP接続ができません"
    Public REGI_ONE_MAX As Integer = 125
    Public REGI_ALL_MAX As Integer = REGI_ONE_MAX * 2
    Public TIME_WRITE_DATA_NUM As Integer = 8   'Modbus書込みデータのデータ数
#End Region

#Region "構造体_Modbus設定値"
    Public Structure sttModbus
        Dim intModTransId As Integer     '転送ID
        Dim intModProtId As Integer      'プロトコルID
        Dim intModUnitId As Integer      'ユニットID
        Dim intModFunc As Integer        'ファンクションコード
        Dim intModReadAddr As Integer    'Read開始アドレス
        Dim intModReadRegist As Integer  'Readレジスタ数
        Dim intModWriteAddr As Integer   'Write開始アドレス
        Dim intModWriteRegist As Integer 'Writeレジスタ数
    End Structure
#End Region

#Region "Gateway処理用"
    Public gtaskGatewayProcessing As Task       'GW通信処理タスク
    Public gblStopFlagGateway As Boolean        '終了フラグ
    Public gblStopTempFlagGateway As Boolean    '中断フラグ
    Public gblthreadsInitialized As Boolean     'Threadの初期チェック
    Public gintDbgCount1GW As Integer           'Thread1のカウント
    Public gintDbgCount2GW As Integer           'Thread2のカウント
    Public gintDbgCount3GW As Integer           'Thread3のカウント
    Public gintDbgCount4GW As Integer           'Thread4のカウント
    Public gintDbgCount5GW As Integer           'Thread5のカウント
    Public gintDbgCount6GW As Integer           'Thread6のカウント
    Public gintDbgCount7GW As Integer           'Thread7のカウント
    Public gintDbgCount8GW As Integer           'Thread8のカウント

    Public gtaskServerProcessing As Task        'サーバ通信処理タスク
    Public gblStopTempFlagServer As Boolean     '終了フラグ
    Public gblStopFlagServer As Boolean         '中断フラグ

    Public WithEvents timerForQueue As System.Timers.Timer  'タイムチェックイベント
    Public timeQueue As Queue(Of String)                    'データ時間保存用Queue

    Public gblCommStartThr1Flag As Boolean          'Thread1用通信開始フラグ
    Public gblCommEndThr1Flag As Boolean            'Thread1用通信終了フラグ
    Public gblCommStartThr2Flag As Boolean          'Thread2用通信開始フラグ
    Public gblCommEndThr2Flag As Boolean            'Thread2用通信終了フラグ
    Public gblCommStartThr3Flag As Boolean          'Thread3用通信開始フラグ
    Public gblCommEndThr3Flag As Boolean            'Thread3用通信終了フラグ
    Public gblCommStartThr4Flag As Boolean          'Thread4用通信開始フラグ
    Public gblCommEndThr4Flag As Boolean            'Thread4用通信終了フラグ
    Public gblCommStartThr5Flag As Boolean          'Thread5用通信開始フラグ
    Public gblCommEndThr5Flag As Boolean            'Thread5用通信終了フラグ
    Public gblCommStartThr6Flag As Boolean          'Thread6用通信開始フラグ
    Public gblCommEndThr6Flag As Boolean            'Thread6用通信終了フラグ
    Public gblCommStartThr7Flag As Boolean          'Thread7用通信開始フラグ
    Public gblCommEndThr7Flag As Boolean            'Thread7用通信終了フラグ
    Public gblCommStartThr8Flag As Boolean          'Thread8用通信開始フラグ
    Public gblCommEndThr8Flag As Boolean            'Thread8用通信終了フラグ

    Public dtTimeTaking As String                   '取得している時間データ
    Public blFormClosing As Boolean
#End Region

#Region "XMLデータ定義"
    Public gwSettingInfo As Integer
    Public serverSettingInfo As Integer
    Public delayTimeInfo As Integer
    Public cycleInfo As Integer

    Public ipAddressInfo As String
    Public userNameInfo As String
    Public passWordInfo As String
    Public pathInfo As String

    Public gw1 As New clsGateway1()
    Public Class clsGateway1
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw2 As New clsGateway2()
    Public Class clsGateway2
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw3 As New clsGateway3()
    Public Class clsGateway3
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw4 As New clsGateway4()
    Public Class clsGateway4
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw5 As New clsGateway5()
    Public Class clsGateway5
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw6 As New clsGateway6()
    Public Class clsGateway6
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw7 As New clsGateway7()
    Public Class clsGateway7
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public gw8 As New clsGateway8()
    Public Class clsGateway8
        Public Property GWName As String
        Public Property IPAddress As String
        Public Property UnitNumber As String
    End Class

    Public arrLoRaAddOfGW1() As Integer
    Public arrModAddOfGW1() As Integer

    Public arrLoRaAddOfGW2() As Integer
    Public arrModAddOfGW2() As Integer

    Public arrLoRaAddOfGW3() As Integer
    Public arrModAddOfGW3() As Integer

    Public arrLoRaAddOfGW4() As Integer
    Public arrModAddOfGW4() As Integer

    Public arrLoRaAddOfGW5() As Integer
    Public arrModAddOfGW5() As Integer

    Public arrLoRaAddOfGW6() As Integer
    Public arrModAddOfGW6() As Integer

    Public arrLoRaAddOfGW7() As Integer
    Public arrModAddOfGW7() As Integer

    Public arrLoRaAddOfGW8() As Integer
    Public arrModAddOfGW8() As Integer
#End Region
End Module
