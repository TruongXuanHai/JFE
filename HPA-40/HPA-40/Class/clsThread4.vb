' HPA-40
' clsThread4.vb
' Threadのタスクを実行
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Threading

Public Class clsThread4
    'Threadsのタスクを実行する
    Public Sub subThread4Processing()
        Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

        While Not gblStopFlagGateway
            If Not gblStopTempFlagGateway Then
                If (gblCommStartThr4Flag = True And gblCommEndThr4Flag = False) Then
                    Thread.Sleep(100)
                    gintDbgCount4GW += 1
                    Console.WriteLine("Count in Thread4: {0}", gintDbgCount4GW)
                    Console.WriteLine("Thread4の時間データ: {0}", dtTimeTaking)
                    '通信開始フラグをFalseにする
                    gblCommStartThr4Flag = False
                    '通信終了フラグをTrueにする
                    gblCommEndThr4Flag = True
                End If
            Else
                Thread.Sleep(500)
                Console.WriteLine("Thread4 (ID:{0})の状態: 待機中", threadId)
            End If
        End While
        Console.WriteLine("Thread4 (ID:{0})の状態: 終了", threadId)
    End Sub
End Class
