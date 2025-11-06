VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BearerToken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mResponse As Dictionary
Private mRequest As WebRequest
Private mStartTime As Single
Private mToken As String
Private mLastsFor As Long
Private mClientID As String
Private mClientSecret As String
Private mAccountID As String
Private Const mURL = "/noapi/https://zoom.us/oauth/token"
Property Get Token() As String
    On Error Resume Next
    If Len(mAccountID) = 0 Or Len(mClientID) = 0 Or Len(mClientSecret) = 0 Then
        frmZoomSecrets.Show
        Class_Initialize
    End If
    If Len(mToken) = 0 Or TimeDiff >= mLastsFor Then
        Dim response As WebResponse, Body As Dictionary
        Set Body = CreateBody(Array("grant_type", "account_id"), Array("account_credentials", mAccountID), False)
        Set response = UseAPI(Body, HttpPost, , , mURL, mClientID, mClientSecret, , FormUrlEncoded)
        mLastsFor = response.Data("expires_in")
        mStartTime = Timer
        mToken = response.Data("access_token")
        Set response = Nothing
    End If
    Token = mToken
End Property
'Property Set Request(Request As WebRequest): Set mRequest = Request: End Property
Property Let ClientID(ClientID As String): mClientID = ClientID: End Property
Property Let ClientSecret(ClientSecret As String): mClientSecret = ClientSecret: End Property
Property Let AccountID(AccountID As String): mAccountID = AccountID: End Property
Private Sub Class_Initialize()
    mAccountID = GetZoomAccountID
    mClientID = GetZoomClientID
    mClientSecret = GetZoomClientSecret
    mStartTime = Timer
End Sub
Private Function TimeDiff() As Long
    TimeDiff = Timer - mStartTime
    If TimeDiff < 0 Then TimeDiff = TimeDiff + 86400
End Function
