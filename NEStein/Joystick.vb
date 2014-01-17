Imports Microsoft.DirectX.DirectInput
Module Joystick
    Dim Joystick_Device(1) As Device
    Dim JoyNum As Integer
    Public Sub Joystick_Initialize()
        'List of attached joysticks
        Dim Joystick_List As DeviceList = Manager.GetDevices(DeviceClass.GameControl, EnumDevicesFlags.AttachedOnly)

        'There is one controller at least?
        JoyNum = Joystick_List.Count
        If JoystickPlugged() Then
            For i As Integer = 0 To JoystickPlugged() - 1
                'Select first joystick
                Joystick_List.MoveNext()

                'Create an object instance
                Dim Dev_Instance As DeviceInstance = Joystick_List.Current

                Joystick_Device(i) = New Device(Dev_Instance.InstanceGuid)
                Joystick_Device(i).SetCooperativeLevel(FrmMain, CooperativeLevelFlags.Background Or CooperativeLevelFlags.NonExclusive)
                Joystick_Device(i).SetDataFormat(DeviceDataFormat.Joystick)
                Joystick_Device(i).Acquire()
            Next
        End If
    End Sub
    Public Function GetJoy_X(ByVal JoyIndex As Integer) As Integer
        If JoystickPlugged() >= (JoyIndex + 1) Then
            Return Joystick_Device(JoyIndex).CurrentJoystickState.X
        Else
            Return Nothing
        End If
    End Function
    Public Function GetJoy_Y(ByVal JoyIndex As Integer) As Integer
        If JoystickPlugged() >= (JoyIndex + 1) Then
            Return Joystick_Device(JoyIndex).CurrentJoystickState.Y
        Else
            Return Nothing
        End If
    End Function
    Public Function GetJoy_Btn(ByVal JoyIndex As Integer, ByVal BtnNum As Byte) As Byte
        If JoystickPlugged() >= (JoyIndex + 1) Then
            Return Joystick_Device(JoyIndex).CurrentJoystickState.GetButtons(BtnNum)
        Else
            Return Nothing
        End If
    End Function
    Public Function JoystickPlugged() As Integer
        Return JoyNum
    End Function
End Module
