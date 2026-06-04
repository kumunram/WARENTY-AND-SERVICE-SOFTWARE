Attribute VB_Name = "PublicScheduler"
Option Explicit

'?? Module ????? Timer ?? ??? ??
Public Sub RunScheduledBackup()
    Application.Run "Perfect_Safety_Module.MasterBackupScheduled"
End Sub

Public Sub AutoSaveTick()
    Application.Run "Perfect_Safety_Module.AutoSaveTick"
End Sub
