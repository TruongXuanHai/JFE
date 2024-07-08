' HPA-40
' clsAllVariable.vb
' グローバル変数
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Public Class clsAllVariable

#Region "ENUM"
#Region "ENUM-受信状態"
    Public Enum RcvStat
        RCV_IDLE = 0      'アイドル
        RCV_MID           '受信中
        RCV_END           '受信完了
        RCV_ERR           'エラー
    End Enum
#End Region
#End Region

End Class



