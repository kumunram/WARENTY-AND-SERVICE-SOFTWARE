VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOTP 
   Caption         =   "Verify OTP"
   ClientHeight    =   1770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5565
   OleObjectBlob   =   "frmOTP.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnVerify_Click()

If Now > otpTime + TimeValue("00:10:00") Then

MsgBox "OTP Expired. Please request new OTP."
Unload Me
Exit Sub

End If

If txtOTP.value = OTPCode Then

MsgBox "OTP Verified Successfully"

frmAddCustomer.txtMobile = frmEntryWizard.txtMobile
frmAddCustomer.Show

Unload Me

Else

MsgBox "Invalid OTP"

End If

End Sub

