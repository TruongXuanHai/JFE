' HPA-40
' clsThread8.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread8
    'Threadsのタスクを実行する
    Public Sub subThread8Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr8Flag = True And gblCommEndThr8Flag = False) Then
                    Thread.Sleep(500)
                    gintDbgCount8GW += 1
                    Console.WriteLine("Count in Thread8: {0}", gintDbgCount8GW)
                    Console.WriteLine("Thread8の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr8Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr8Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread8 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread8 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
