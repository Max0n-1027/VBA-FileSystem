VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("RyoYoshizawa.FileSystem")
Option Explicit

Private stream As Object
Private utf8NoBOM As Boolean

Public Enum adCharSetEnum
    ASCII = 20127
    ISO_8859_1 = 28591
    Shift_Jis = 932
    UTF_7 = 65000
    UTF_8 = 65001
    UTF_8n = -65001
    UTF_16LE = 1200
    UTF_16BE = 1201
End Enum

Public Enum adConnectModeEnum
    adModeUnknown = 0
    adModeRead = 1
    adModeWrite = 2
    adModeReadWrite = 3
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeShareDenyNone = 16
    adModeRecursive = 4194304
End Enum

Public Enum adStreamTypeEnum
    adTypeBinary = 1
    adTypeText = 2
End Enum

Public Enum adStateEnum
    adStateClosed = 0
    adStateOpen = 1
    adStateConnecting = 2
    adStateExecuting = 4
    adStateFetching = 8
End Enum

Public Enum adStreamWriteEnum
    adWriteChar = 0
    adWriteLine = 1
End Enum

Public Enum adSaveOptionsEnum
    adSaveCreateNotExist = 1
    adSaveCreateOverWrite = 2
End Enum

Public Enum adLineSeparatorsEnum
    adCRLF = -1
    adLF = 10
    adCR = 13
End Enum

Public Enum adStreamOpenOptionsEnum
    adOpenStreamUnspecified = -1
    adOpenStreamAsync = 1
    adOpenStreamFromRecord = 4
End Enum

Public Property Get Charset() As String
    Charset = stream.Charset
End Property

Public Property Get DataType() As adStreamTypeEnum
    DataType = stream.Type
End Property

Public Property Let DataType(ByVal dType As adStreamTypeEnum)
    stream.Type = dType
End Property

Public Function EOS() As Boolean
    EOS = stream.EOS
End Function

Public Property Get LineSeparator() As adLineSeparatorsEnum
    LineSeparator = stream.LineSeparator
End Property

Public Property Let LineSeparator(ByVal SeparatorType As adLineSeparatorsEnum)
    stream.LineSeparator = SeparatorType
End Property

Public Property Get Mode() As adConnectModeEnum
    Mode = stream.Mode
End Property

Public Property Let Mode(ByVal Mode As adConnectModeEnum)
    stream.Mode = Mode
End Property

Public Property Get Position() As Long
    Position = stream.Position
End Property

Public Property Let Position(ByVal pos As Long)
    stream.Position = pos
End Property

Public Property Get Size() As Long
    Size = stream.Size
End Property

Private Sub Class_Initialize()
    Set stream = CreateObject("ADODB.Stream")
End Sub

Private Sub Class_Terminate()
    If stream.State <> adStateEnum.adStateClosed Then
        stream.Close
    End If
End Sub

Public Function Cancel() As Boolean
    On Error GoTo Catch
    stream.Cancel
    Cancel = True
    Exit Function
Catch:
    Cancel = False
End Function

Public Function CloseStream() As Boolean
    On Error GoTo Catch
    stream.Close
    CloseStream = True
    Exit Function
Catch:
    CloseStream = False
End Function

Public Function CopyTo(ByRef DestStream As Object, Optional NumChars As Integer = -1)
    On Error GoTo Catch
    stream.CopyTo DestStream, NumChars
    CopyTo = True
    Exit Function
Catch:
    CopyTo = False
End Function

Public Function Flush() As Boolean
    On Error GoTo Catch
    stream.Flush
    Flush = True
    Exit Function
Catch:
    Flush = False
End Function

Public Function LoadFromFile(ByVal FileName As String) As Boolean
    On Error GoTo Catch
    stream.LoadFromFile FileName
    LoadFromFile = True
    Exit Function
Catch:
    LoadFromFile = False
End Function

Public Function OpenStream( _
       Optional ByVal Source As Variant, _
       Optional ByVal Mode As adConnectModeEnum = adConnectModeEnum.adModeUnknown, _
       Optional ByVal OpenOptions As adStreamOpenOptionsEnum = adStreamOpenOptionsEnum.adOpenStreamUnspecified, _
       Optional ByVal UserName As String, _
       Optional ByVal Password As String) As Boolean
    On Error GoTo Catch
    stream.Open Source, Mode, OpenOptions, UserName, Password
    OpenStream = True
    Exit Function
Catch:
    OpenStream = False
End Function

Public Function ReadBinary(Optional NumBytes As Long = -1) As Variant
    ReadBinary = stream.Read(NumBytes)
End Function

Public Function ReadText(Optional ByVal NumChars As Long = -1) As String
    ReadText = stream.ReadText(NumChars)
End Function

Public Function SaveToFile( _
       ByVal FileName As String, _
       Optional ByVal SaveOptions As adSaveOptionsEnum = adSaveOptionsEnum.adSaveCreateNotExist _
       ) As Boolean
    On Error GoTo Catch
    stream.SaveToFile FileName, SaveOptions
    If utf8NoBOM Then
        Dim bt As Variant
        Dim st As Object: Set st = CreateObject("ADODB.Stream")
        st.Type = adTypeBinary
        st.Open
        st.LoadFromFile FileName
        st.Position = 3
        bt = st.Read
        st.Position = 0
        st.Write bt
        st.SetEOS
        st.SaveToFile FileName, adSaveOptionsEnum.adSaveCreateOverWrite
        st.Close
        Set st = Nothing
    End If
    SaveToFile = True
    Exit Function
Catch:
    SaveToFile = False
End Function

Public Function SetCharset(ByVal char As adCharSetEnum)
    Select Case char
    Case adCharSetEnum.ASCII
        stream.Charset = "ascii"
    Case adCharSetEnum.ISO_8859_1
        stream.Charset = "iso-8859-1"
    Case adCharSetEnum.Shift_Jis
        stream.Charset = "shift_jis"
    Case adCharSetEnum.UTF_7
        stream.Charset = "utf-7"
    Case adCharSetEnum.UTF_8, adCharSetEnum.UTF_8n
        stream.Charset = "utf-8"
    Case adCharSetEnum.UTF_16LE
        stream.Charset = "unicode"
    Case adCharSetEnum.UTF_16BE
        stream.Charset = "unicodeFFFE"
    End Select
    If char = UTF_8n Then
        utf8NoBOM = True
    Else
        utf8NoBOM = False
    End If
End Function

Function SetEOS() As Boolean
    On Error GoTo Catch
    stream.SetEOS
    SetEOS = True
    Exit Function
Catch:
    SetEOS = False
End Function

Public Function SkipLine() As Boolean
    On Error GoTo Catch
    stream.SkipLine
    SkipLine = True
    Exit Function
Catch:
    SkipLine = False
End Function

Public Function WriteBinary(ByVal Buffer As Variant) As Boolean
    On Error GoTo Catch
    stream.Write Buffer
    WriteBinary = True
    Exit Function
Catch:
    WriteBinary = False
End Function

Public Function WriteText( _
       ByVal Data As String, _
       Optional Options As adStreamWriteEnum = adStreamWriteEnum.adWriteChar _
       ) As Boolean
    On Error GoTo Catch
    stream.WriteText Data, Options
    WriteText = True
    Exit Function
Catch:
    WriteText = False
End Function


