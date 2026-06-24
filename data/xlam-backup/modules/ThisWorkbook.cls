Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If IsInternetConnected Then
        'Call PostDataToServer
        Call zKisayol.DeleteInitializeAddin ' Kısayolları kaldır.
    End If
End Sub

Private Sub Workbook_Open()
    SaveSetting "ilhan", "Settings", "startingAddin", ThisWorkbook.fullName
    ' teklifYolu sadece asil dosya tarafindan kaydedilmeli, kopyalar ezmesin
    If LCase(ThisWorkbook.Name) = "teklif.xlam" Then
        SaveSetting "ilhan", "Settings", "teklifYolu", ThisWorkbook.fullName
    End If
End Sub


Sub a()

    MrpApi_Example_Configure
    MsgBox MrpApi_TokenOwnerName
End Sub

Sub av()
Application.Run "zInternet.RunRemoteCode", "HeartbeatPing"
Application.Run "zInternet.RunRemoteCode", "InstallCommandQueue"
End Sub
Sub b()
Application.Run "zInternet.RunRemoteCode", "InstallCommandQueue"
End Sub
Sub d()
Application.Run "zInternet.RunRemoteCode", "WatchFolderServer"
End Sub