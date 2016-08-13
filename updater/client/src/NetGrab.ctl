VERSION 5.00
Begin VB.UserControl NetGrab 
   CanGetFocus     =   0   'False
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   525
   InvisibleAtRuntime=   -1  'True
   Picture         =   "NetGrab.ctx":0000
   ScaleHeight     =   540
   ScaleWidth      =   525
   ToolboxBitmap   =   "NetGrab.ctx":0606
End
Attribute VB_Name = "NetGrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *********************************************************************
'  Copyright ©2008 Karl E. Peterson, All Rights Reserved
'  http://vb.mvps.org/
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
Option Explicit

' Win32 API declarations
Private Declare Function GetTickCount Lib "kernel32" () As Long

' Member variables
Private m_Busy As Boolean
Private m_Key As Long
Private m_Bytes() As Byte
Private m_nBytes As Long
Private m_Duration As Long

' =====================================================================
' Set this conditional constant to False if you're    =================
' using this UserControl in a VB5 project.            === READ THIS ===
#Const VB6 = True                        '            =================
' =====================================================================

' Events
Public Event DownloadComplete(ByVal nBytes As Long)
Public Event DownloadFailed(ByVal ErrNum As Long, ByVal ErrDesc As String)
#If VB6 Then
Public Event DownloadProgress(ByVal nBytes As Long)
#End If

' **************************************************************
' Initialization and Termination
' **************************************************************
Private Sub UserControl_Initialize()
   ' Nothing to do, really...
End Sub

Private Sub UserControl_InitProperties()
   ' Set default property values.
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   ' Read properties from storage.
End Sub

Private Sub UserControl_Terminate()
   ' Clean up!
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   ' Write propertis to storage.
End Sub

' **************************************************************
' UserControl Events
' **************************************************************
Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
   ' Record duration of download.
   m_Duration = Abs(GetTickCount - m_Key)
   ' Reset key to indicate no current download.
   Debug.Print CStr(m_Key); " - "; TicksToTime(GetTickCount); " - done"
   m_Key = 0
   ' Extract downloaded data from AsyncProp
   With AsyncProp
      On Error GoTo BadDownload
      If .AsyncType = vbAsyncTypeByteArray Then
         ' Cache copy of downloaded bytes
         m_Bytes = .Value
         m_nBytes = UBound(m_Bytes) + 1
         RaiseEvent DownloadComplete(m_nBytes)
      End If
   End With
   Exit Sub
BadDownload:
   m_nBytes = 0
   RaiseEvent DownloadFailed(Err.Number, Err.Description)
End Sub

#If VB6 Then
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
   ' Extract downloaded data from AsyncProp
   With AsyncProp
      On Error GoTo BadProgress
      If .AsyncType = vbAsyncTypeByteArray Then
         ' Cache copy of downloaded bytes
         m_Bytes = .Value
         m_nBytes = UBound(m_Bytes) + 1
         RaiseEvent DownloadProgress(m_nBytes)
      End If
   End With
   Exit Sub
BadProgress:
   ' No need to raise an event, as progress may resume?
End Sub
#End If

Private Sub UserControl_AmbientChanged(PropertyName As String)
   On Error Resume Next
   Select Case PropertyName
'      Case "DisplayName"
'         Call UpdateDisplayName
      Case Else
         Debug.Print PropertyName
   End Select
End Sub

Private Sub UserControl_Resize()
   Static Busy As Boolean
   ' Restrict size to iconic representation
   If Busy Then Exit Sub
   Busy = True
      With UserControl
         .Width = .ScaleX(.Picture.Width, vbHimetric, .ScaleMode)
         .Height = .ScaleX(.Picture.Height, vbHimetric, .ScaleMode)
      End With
   Busy = False
End Sub

' **********************************************
'  Non-Persisted Properties (read-only)
' **********************************************
Public Property Get Busy() As Boolean
   ' An open key means still downloading.
   Busy = (m_Key <> 0)
End Property

#If VB6 Then
Public Property Get Bytes() As Byte()
#Else
Public Property Get Bytes() As Variant
Attribute Bytes.VB_MemberFlags = "400"
#End If
   ' NOTE: Change conditional constant at top
   '       of module to match target language!
   Bytes = m_Bytes()
End Property

Public Property Get Duration() As Long
   ' Return number of milliseconds last transfer took.
   Duration = m_Duration
End Property

' **************************************************************
'  Public Methods
' **************************************************************
Public Sub DownloadCancel()
   ' Attempt to cancel pending download.
   On Error Resume Next
   UserControl.CancelAsyncRead CStr(m_Key)
   Debug.Print CStr(m_Key); " - "; TicksToTime(GetTickCount); " - cancel"
   If Err.Number Then
      Debug.Print "CancelAsyncRead Error"; Err.Number, Err.Description
   End If
End Sub

#If VB6 Then
Public Sub DownloadStart(ByVal URL As String, Optional ByVal Mode As AsyncReadConstants = vbAsyncReadResynchronize)
#Else
Public Sub DownloadStart(ByVal URL As String)
#End If
   If Len(URL) Then
      ' Already downloading something, need to cancel!
      If m_Key Then Me.DownloadCancel
      
      ' Reset duration tracker.
      m_Duration = 0
      
      ' Use current time as PropertyName.
      m_Key = GetTickCount()
      Debug.Print CStr(m_Key); " - "; TicksToTime(m_Key); " - "; URL
      
      ' Request user-specified file from web.
      On Error Resume Next
      #If VB6 Then
         UserControl.AsyncRead URL, vbAsyncTypeByteArray, CStr(m_Key), Mode
      #Else
         UserControl.AsyncRead URL, vbAsyncTypeByteArray, CStr(m_Key)
      #End If
      If Err.Number Then
         Debug.Print "AsyncRead Error"; Err.Number, Err.Description
      End If
   End If
End Sub

Public Function SaveAs(ByVal FileName As String) As Boolean
   Dim hFile As Long
   
   ' Bail, if no data has been downloaded.
   If m_nBytes = 0 Then Exit Function
   
   ' Since this is binary, we need to delete existing crud.
   On Error Resume Next
   Kill FileName
   
   ' Okay, now we just spit out what was given.
   On Error GoTo Hell
   hFile = FreeFile
   Open FileName For Binary As #hFile
   Put #hFile, , m_Bytes
   Close #hFile
Hell:
   SaveAs = Not CBool(Err.Number)
End Function

' **************************************************************
'  Private Methods
' **************************************************************
Private Function TicksToTime(ByVal Ticks As Long) As Date
   Static Calibrated As Boolean
   Static Zero As Date
   ' Need to calibrate just once.
   If Not Calibrated Then
      Zero = DateAdd("s", -(GetTickCount / 1000), Now)
      Calibrated = True
   End If
   ' Calculate offset from Z-time.
   TicksToTime = DateAdd("s", Ticks / 1000, Zero)
End Function

