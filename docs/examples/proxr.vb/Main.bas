Attribute VB_Name = "MainModule"
Sub Main()
    frmSelectDevice.Show vbModal
    If Not frmSelectDevice.SelectOKButton Then Exit Sub
    Dim s As String
    s = GetSetting("ProXR_V4", "Settings", "DeviceType", 0)
    Dim i As Integer
    i = Val(s)
    Select Case i
        Case MostNcdDevicesR2X
            ShowMostNcdDevicesR2X
        Case MostNcdDevicesR8X
            ShowMostNcdDevicesR8X
        Case MostNcdDevicesProXr
            ShowMostNcdDevicesProXr
        Case ZigBeeS1ZR2X
            ShowZigBeeS1ZR2X
        Case ZigBeeS1ProXr
            ShowZigBeeS1ProXr
        Case ZigBeeMeshR2X
            ShowZigBeeMeshR2X
        Case ZigBeeMeshProXr
            ShowZigBeeMeshProXr
        Case XscLongRangeR2X
            ShowXscLongRangeR2X
        Case XscLongRangeProXr
            ShowXscLongRangeProXr
    End Select
    Unload frmSelectDevice
    
End Sub

Sub ShowMostNcdDevicesR2X()
    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR2X.ProXR1.BaudRate = 9600
    frmR2X.ProXR1.OpenPort
    frmR2X.Show
    'Unload frmR2X
    'Debug.Print "ShowMostNcdDevicesR2X"
End Sub

Sub ShowMostNcdDevicesR8X()
    frmR8X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR8X.ProXR1.BaudRate = Val(frmSelectDevice.cmbBaudrate.Text)
    frmR8X.ProXR1.OpenPort
    frmR8X.Show
    'Unload frmR8X
End Sub

Sub ShowMostNcdDevicesProXr()
    Form1.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    Form1.ProXR1.BaudRate = br_115200
    Form1.ProXR1.OpenPort
    Form1.LoadForm
    Form1.Show
    'Unload Form1
End Sub

Sub ShowZigBeeS1ZR2X()
    frmSelectZigbeeS1Device.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeS1Device.ProXR1.BaudRate = br_115200
    frmSelectZigbeeS1Device.ProXR1.OpenPort
    frmSelectZigbeeS1Device.Show vbModal
    frmSelectZigbeeS1Device.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectZigbeeS1Device.SelectOKButton
    Unload frmSelectZigbeeS1Device
    If Not selectOK Then Exit Sub
    
    
    
    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR2X.ProXR1.BaudRate = 115200
    frmR2X.ProXR1.OpenPort
    frmR2X.Show
    Unload frmR2X
End Sub

Sub ShowZigBeeS1ProXr()
    frmSelectZigbeeS1Device.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeS1Device.ProXR1.BaudRate = br_115200
    frmSelectZigbeeS1Device.ProXR1.OpenPort
    frmSelectZigbeeS1Device.Show vbModal
    frmSelectZigbeeS1Device.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectZigbeeS1Device.SelectOKButton
    Unload frmSelectZigbeeS1Device
    If Not selectOK Then Exit Sub
    
    
    ShowMostNcdDevicesProXr
End Sub

Sub ShowZigBeeMeshR2X()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = br_115200
    frmSelectZigbeeDevice.ProXR1.OpenPort
    frmSelectZigbeeDevice.Show vbModal
    frmSelectZigbeeDevice.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectZigbeeDevice.SelectOKButton
    Unload frmSelectZigbeeDevice
    If Not selectOK Then Exit Sub
    
    frmR2X.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmR2X.ProXR1.BaudRate = 115200
    frmR2X.ProXR1.OpenPort
    frmR2X.Show
    Unload frmR2X
End Sub

Sub ShowZigBeeMeshProXr()
    frmSelectZigbeeDevice.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectZigbeeDevice.ProXR1.BaudRate = br_115200
    frmSelectZigbeeDevice.ProXR1.OpenPort
    frmSelectZigbeeDevice.Show vbModal
    frmSelectZigbeeDevice.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectZigbeeDevice.SelectOKButton
    Unload frmSelectZigbeeDevice
    If Not selectOK Then Exit Sub
    
    
    Form1.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    Form1.ProXR1.BaudRate = br_115200
    Form1.ProXR1.OpenPort
    Form1.LoadForm
    Form1.Show
End Sub

Sub ShowXscLongRangeR2X()
    frmSelectXSC.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectXSC.ProXR1.BaudRate = br_115200
    frmSelectXSC.ProXR1.OpenPort
    frmSelectXSC.Show vbModal
    frmSelectXSC.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectXSC.SelectOKButton
    Unload frmSelectXSC
    If Not selectOK Then Exit Sub
    ShowZigBeeS1ZR2X

End Sub
        
Sub ShowXscLongRangeProXr()
    frmSelectXSC.ProXR1.PortName = frmSelectDevice.cmbPorts.Text
    frmSelectXSC.ProXR1.BaudRate = br_115200
    frmSelectXSC.ProXR1.OpenPort
    frmSelectXSC.Show vbModal
    frmSelectXSC.ProXR1.ClosePort
    Dim selectOK As Boolean
    selectOK = frmSelectXSC.SelectOKButton
    Unload frmSelectXSC
    If Not selectOK Then Exit Sub
    ShowZigBeeS1ZR2X
End Sub

