' HPA-40
' clsThread6.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread7
    'Threadsのタスクを実行する
    Public Sub subThread7Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr7Flag = True And gblCommEndThr7Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount7GW += 1
                    Console.WriteLine("Count in Thread7: {0}", gintDbgCount7GW)
                    Console.WriteLine("Thread7の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr7Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr7Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread7 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread7 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
