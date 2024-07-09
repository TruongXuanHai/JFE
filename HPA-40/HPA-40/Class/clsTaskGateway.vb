' HPA-40
' getXmlData.vb
' XMLファイルから設定値を取得
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading.Tasks
Imports System.Threading
Imports System.Runtime.InteropServices
Imports System.Timers
Imports System.IO

Public Class clsTaskGateway
#Region "クラスを宣言"
    Public clsThread1 As New clsThread1() 'Thread 1
    Public clsThread2 As New clsThread2() 'Thread 2
    Public clsThread3 As New clsThread3() 'Thread 3
    Public clsThread4 As New clsThread4() 'Thread 4
    Public clsThread5 As New clsThread5() 'Thread 5
    Public clsThread6 As New clsThread6() 'Thread 6
    Public clsThread7 As New clsThread7() 'Thread 7
    Public clsThread8 As New clsThread8() 'Thread 8
#End Region

#Region "スレッドを宣言"
    Public thread1 As Thread
    Public thread2 As Thread
    Public thread3 As Thread
    Public thread4 As Thread
    Public thread5 As Thread
    Public thread6 As Thread
    Public thread7 As Thread
    Public thread8 As Thread
#End Region

#Region "関数_Task1を作成"
    Public Sub subInitializeTaskGateway()
        'タスクの存在を確認
        If gtaskGatewayProcessing Is Nothing Then
            gtaskGatewayProcessing = New Task(Sub() subRunTaskGateway())
            gtaskGatewayProcessing.Start()

            '開始/終了通信チェック用フラグをFalseにする
            gblCommStartThr1Flag = False
            gblCommEndThr1Flag = False
            gblCommStartThr2Flag = False
            gblCommEndThr2Flag = False
            gblCommStartThr3Flag = False
            gblCommEndThr3Flag = False
            gblCommStartThr4Flag = False
            gblCommEndThr4Flag = False
            gblCommStartThr5Flag = False
            gblCommEndThr5Flag = False
            gblCommStartThr6Flag = False
            gblCommEndThr6Flag = False
            gblCommStartThr7Flag = False
            gblCommEndThr7Flag = False
            gblCommStartThr8Flag = False
            gblCommEndThr8Flag = False

            'カウンター初期化
            gintDbgCount1GW = 0
            gintDbgCount2GW = 0
            gintDbgCount3GW = 0
            gintDbgCount4GW = 0
            gintDbgCount5GW = 0
            gintDbgCount6GW = 0
            gintDbgCount7GW = 0
            gintDbgCount8GW = 0
        End If
    End Sub
#End Region

#Region "関数_Task1を実行"
    Private Sub subRunTaskGateway()
        If Not gblthreadsInitialized Then
            'Threadsを作成
            thread1 = New Thread(Sub() clsThread1.subThread1Processing())
            thread2 = New Thread(Sub() clsThread2.subThread2Processing())
            thread3 = New Thread(Sub() clsThread3.subThread3Processing())
            thread4 = New Thread(Sub() clsThread4.subThread4Processing())
            thread5 = New Thread(Sub() clsThread5.subThread5Processing())
            thread6 = New Thread(Sub() clsThread6.subThread6Processing())
            thread7 = New Thread(Sub() clsThread7.subThread7Processing())
            thread8 = New Thread(Sub() clsThread8.subThread8Processing())
            'Threadsを起動
            thread1.Start()
            thread2.Start()
            thread3.Start()
            thread4.Start()
            thread5.Start()
            thread6.Start()
            thread7.Start()
            thread8.Start()
            'Threadの初期フラグをTrueにする
            gblthreadsInitialized = True
        End If

        Dim blcheckRequireFirstTime As Boolean = True
        While Not gblStopFlagGateway
            If gblStopTempFlagGateway Then
                Thread.Sleep(500)
                Console.WriteLine("Task1の状態: 待機中")
            Else
                If timeQueue.Count <> 0 Then
                    'Console.WriteLine("最初の時間データ： " & timeQueue.Peek())
                    If blcheckRequireFirstTime Then
                        'Queueから次の時間データを取る
                        dtTimeTaking = timeQueue.Dequeue().ToString()
                        Console.WriteLine("取得している時間データ： " & dtTimeTaking)

                        '開始させるためにフラグを設定
                        subFlagSettingForStart()

                        blcheckRequireFirstTime = False
                    End If

                    '全てのThreadが完了
                    If gblCommEndThr1Flag = True And gblCommEndThr2Flag = True And _
                        gblCommEndThr3Flag = True And gblCommEndThr4Flag = True And _
                        gblCommEndThr5Flag = True And gblCommEndThr6Flag = True And _
                        gblCommEndThr7Flag = True And gblCommEndThr8Flag = True Then

                        'Queueから次の時間データを取る
                        dtTimeTaking = timeQueue.Dequeue()
                        Console.WriteLine("取得している時間データ： " & dtTimeTaking)

                        '改めて開始させるためにフラグを設定
                        subFlagSettingForStart()
                    End If
                Else
                    blcheckRequireFirstTime = True
                End If
            End If

        End While
        Console.WriteLine("Task1の状態: 終了")
    End Sub
#End Region

#Region "関数_すべてのThreadが開始するために、フラグを設定"
    Private Sub subFlagSettingForStart()
        '通信開始フラグをTrueにする
        gblCommStartThr1Flag = True
        gblCommStartThr2Flag = True
        gblCommStartThr3Flag = True
        gblCommStartThr4Flag = True
        gblCommStartThr5Flag = True
        gblCommStartThr6Flag = True
        gblCommStartThr7Flag = True
        gblCommStartThr8Flag = True

        '通信終了フラグをFalseにする
        gblCommEndThr1Flag = False
        gblCommEndThr2Flag = False
        gblCommEndThr3Flag = False
        gblCommEndThr4Flag = False
        gblCommEndThr5Flag = False
        gblCommEndThr6Flag = False
        gblCommEndThr7Flag = False
        gblCommEndThr8Flag = False
    End Sub
#End Region

#Region "関数_すべてのThreadを強制的に終了"
    Public Sub subStopThreads()
        For Each threadTemp In {thread1, thread2, thread3, thread4, thread5, thread6, thread7, thread8}
            If threadTemp IsNot Nothing Then
                'Threadは10秒以内に終了できない場合、エラー扱いにする
                'Thread.Join(Int32)→スレッドが終了した場合は戻り値：True
                'millisecondsTimeoutパラメーターで指定した時間が経過してもスレッドが終了していない場合は戻り値：False
                '10s = 10000ms
                If Not threadTemp.Join(10000) Then
                    '強制的に終了する
                    threadTemp.Abort()
                    Console.WriteLine("Thread {0} 強制的に終了します。", threadTemp.ManagedThreadId)
                End If
            End If
        Next
    End Sub
#End Region
End Class
