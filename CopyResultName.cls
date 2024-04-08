VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CopyResultName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Implements LSEXT.IEntityExtension

Option Explicit

 

Private aConnection As New ADODB.Connection
Private NtlsCon As LSSERVICEPROVIDERLib.NautilusDBConnection
Private sp As LSSERVICEPROVIDERLib.NautilusServiceProvider
Private Function IEntityExtension_CanExecute(Parameters As LSEXT.IExtensionParameters) As LSEXT.ExecuteExtension

    'Default to show up on the menu
    IEntityExtension_CanExecute = exEnabled

End Function
Private Sub IEntityExtension_Execute(Parameters As LSEXT.IExtensionParametersEx)
    Dim RecordRs As ADODB.Recordset
    Dim aDatasource As String
    Dim strSQL As String

    Set sp = Parameters("SERVICE_PROVIDER")

    Set NtlsCon = sp.QueryServiceProvider("DBConnection")
    aDatasource = NtlsCon.GetADOConnectionString
        aDatasource = "Provider=OraOLEDB.Oracle" & _
              ";Data Source=" & NtlsCon.GetServerDetails & _
              ";User ID=" & NtlsCon.GetUsername & _
              ";Password=" & NtlsCon.GetPassword
 
          If NtlsCon.GetServerIsProxy Then
            aDatasource = "Provider=OraOLEDB.Oracle;Data Source=" & _
            NtlsCon.GetServerDetails & ";User id=/;Persist Security Info=True;"
          End If
    aConnection.Open aDatasource
    aConnection.CursorLocation = adUseClient

    Set RecordRs = Parameters("RECORDS")

    If Not RecordRs.EOF Then
        Call CopyEntityNameToClipbord("Result", RecordRs, aConnection)
    End If
 
    aConnection.Close
End Sub


