' HPA-40
' clsThread6.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread6
    'Threadsのタスクを実行する
    Public Sub subThread6Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr6Flag = True And gblCommEndThr6Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount6GW += 1
                    Console.WriteLine("Count in Thread6: {0}", gintDbgCount6GW)
                    Console.WriteLine("Thread6の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr6Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr6Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread6 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread6 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
