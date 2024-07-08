' HPA-40
' clsGetParameter.vb
' グローバル変数
'
'CORYRIGHT(C) 2024 HAKARU PLUS CORPORATION
'
' 修正履歴
' 2024/06/17 チュオンスアンハイ

Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO

Public Class clsGetParameter
#Region "XMLの構成を構築"
    Public settings As UnitSettings
    <XmlRoot("settings")>
    Public Structure UnitSettings
        <XmlElement("GWNumber")>
        Public Property GWNumber As Integer
        <XmlElement("GWSetting")>
        Public Property GWSetting As Integer
        <XmlElement("ServerSetting")>
        Public Property ServerSetting As Integer
        <XmlElement("DelayTime")>
        Public Property DelayTime As Integer
        <XmlElement("Cycle")>
        Public Property Cycle As Integer
        <XmlElement("GW")>
        Public Property Gateways As List(Of Gateway)
    End Structure

    Public gateways() As Gateway
    Public Structure Gateway
        <XmlElement("GWName")>
        Public Property GWName As String
        <XmlElement("IPAddress")>
        Public Property IPAddress As String
        <XmlElement("UnitNumber")>
        Public Property UnitNumber As Integer
        <XmlElement("Unit")>
        Public Property Units As List(Of Unit)
    End Structure

    Public units() As Unit
    Public Structure Unit
        Public Property UnitType As String
        <XmlElement("LoRaAddress")>
        Public Property LoRaAddress As String
        <XmlElement("RS485Address")>
        Public Property RS485Address As String
    End Structure

    Public svSettings As ServerSettings
    <XmlRoot("settings")>
    Public Structure ServerSettings
        <XmlElement("IPAddress")>
        Public Property IPAddress As String
        <XmlElement("Username")>
        Public Property Username As String
        <XmlElement("Password")>
        Public Property Password As String
        <XmlElement("Path")>
        Public Property Path As String
    End Structure
#End Region

#Region "XMLファイルを読み取って変数を戻る"
    Public Sub subLoadXMLSetting(ByVal fileTempPathGW As String, ByVal fileTempPathServer As String)
        'XMLファイル（Unitsetting.xml）を読み込み
        Dim serialGW As New XmlSerializer(GetType(UnitSettings))
        Using readerGW As New StreamReader(fileTempPathGW)
            settings = CType(serialGW.Deserialize(readerGW), UnitSettings)
        End Using

        'XMLファイル（Serversetting.xml）を読み込み
        Dim serialServer As New XmlSerializer(GetType(ServerSettings))
        Using readerServer As New StreamReader(fileTempPathServer)
            svSettings = CType(serialServer.Deserialize(readerServer), ServerSettings)
        End Using

        '変数に格納する
        gwSettingInfo = settings.GWSetting
        serverSettingInfo = settings.ServerSetting
        delayTimeInfo = settings.DelayTime
        cycleInfo = settings.Cycle

        ipAddressInfo = svSettings.IPAddress
        userNameInfo = svSettings.Username
        passWordInfo = svSettings.Password
        pathInfo = svSettings.Path

        ReDim gateways(settings.Gateways.Count - 1)
        For indexGW As Integer = 0 To settings.Gateways.Count - 1
            gateways(indexGW) = settings.Gateways(indexGW)
        Next

        gw1.GWName = gateways(0).GWName
        gw1.IPAddress = gateways(0).IPAddress
        gw1.UnitNumber = gateways(0).UnitNumber

        'gw2.GWName = gateways(1).GWName
        'gw2.IPAddress = gateways(1).IPAddress
        'gw2.UnitNumber = gateways(1).UnitNumber

        'gw3.GWName = gateways(2).GWName
        'gw3.IPAddress = gateways(2).IPAddress

        'gw4.GWName = gateways(3).GWName
        'gw4.IPAddress = gateways(3).IPAddress

        'gw5.GWName = gateways(4).GWName
        'gw5.IPAddress = gateways(4).IPAddress

        'gw6.GWName = gateways(5).GWName
        'gw6.IPAddress = gateways(5).IPAddress

        'gw7.GWName = gateways(6).GWName
        'gw7.IPAddress = gateways(6).IPAddress

        'gw8.GWName = gateways(7).GWName
        'gw8.IPAddress = gateways(7).IPAddress

        Dim loRaAddTemp As Integer
        Dim modbusAddTemp As Integer

        ReDim arrLoRaAddOfGW1(gateways(0).UnitNumber - 1)
        ReDim arrModAddOfGW1(gateways(0).UnitNumber - 1)
        For indexUnit As Integer = 0 To gateways(0).UnitNumber - 1
            loRaAddTemp = gateways(0).Units(indexUnit).LoRaAddress
            modbusAddTemp = gateways(0).Units(indexUnit).RS485Address
            arrLoRaAddOfGW1(indexUnit) = loRaAddTemp
            arrModAddOfGW1(indexUnit) = modbusAddTemp
        Next

        'ReDim arrLoRaAddOfGW2(gateways(1).UnitNumber - 1)
        'ReDim arrModAddOfGW2(gateways(1).UnitNumber - 1)
        'For indexUnit As Integer = 0 To gateways(1).UnitNumber - 1
        '    loRaAddTemp = gateways(1).Units(indexUnit).LoRaAddress
        '    modbusAddTemp = gateways(1).Units(indexUnit).RS485Address
        '    arrLoRaAddOfGW2(indexUnit) = loRaAddTemp
        '    arrModAddOfGW2(indexUnit) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW3(gateways(2).UnitNumber - 1)
        'ReDim arrModAddOfGW3(gateways(2).UnitNumber - 1)
        'For i As Integer = 0 To gateways(2).UnitNumber - 1
        '    loRaAddTemp = gateways(2).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(2).Units(i).RS485Address
        '    arrLoRaAddOfGW3(2) = loRaAddTemp
        '    arrModAddOfGW3(2) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW4(gateways(3).UnitNumber - 1)
        'ReDim arrModAddOfGW4(gateways(3).UnitNumber - 1)
        'For i As Integer = 0 To gateways(3).UnitNumber - 1
        '    loRaAddTemp = gateways(3).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(3).Units(i).RS485Address
        '    arrLoRaAddOfGW4(i) = loRaAddTemp
        '    arrModAddOfGW4(i) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW5(gateways(4).UnitNumber - 1)
        'ReDim arrModAddOfGW5(gateways(4).UnitNumber - 1)
        'For i As Integer = 0 To gateways(4).UnitNumber - 1
        '    loRaAddTemp = gateways(4).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(4).Units(i).RS485Address
        '    arrLoRaAddOfGW5(i) = loRaAddTemp
        '    arrModAddOfGW5(i) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW6(gateways(5).UnitNumber - 1)
        'ReDim arrModAddOfGW6(gateways(5).UnitNumber - 1)
        'For i As Integer = 0 To gateways(5).UnitNumber - 1
        '    loRaAddTemp = gateways(5).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(5).Units(i).RS485Address
        '    arrLoRaAddOfGW6(i) = loRaAddTemp
        '    arrModAddOfGW6(i) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW7(gateways(6).UnitNumber - 1)
        'ReDim arrModAddOfGW7(gateways(6).UnitNumber - 1)
        'For i As Integer = 0 To gateways(6).UnitNumber - 1
        '    loRaAddTemp = gateways(6).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(6).Units(i).RS485Address
        '    arrLoRaAddOfGW7(i) = loRaAddTemp
        '    arrModAddOfGW7(i) = modbusAddTemp
        'Next

        'ReDim arrLoRaAddOfGW8(gateways(7).UnitNumber - 1)
        'ReDim arrModAddOfGW8(gateways(7).UnitNumber - 1)
        'For i As Integer = 0 To gateways(7).UnitNumber - 1
        '    loRaAddTemp = gateways(7).Units(i).LoRaAddress
        '    modbusAddTemp = gateways(7).Units(i).RS485Address
        '    arrLoRaAddOfGW8(i) = loRaAddTemp
        '    arrModAddOfGW8(i) = modbusAddTemp
        'Next
    End Sub
#End Region

#Region "XMLファイルにGW通信状態とサーバ通信状態を書き込む"
    Public Sub subWriteXMLSetting(ByVal fileTempPathGW As String)
        Dim serialGW As New XmlSerializer(GetType(UnitSettings))
        Using writerGW As New StreamWriter(fileTempPathGW)
            serialGW.Serialize(writerGW, settings)
        End Using
    End Sub
#End Region
End Class
