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

Private Const Stream_DEBUG As Boolean = False
Private Stream As Object
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
    Charset = Stream.Charset
End Property

Public Property Get DataType() As adStreamTypeEnum
    DataType = Stream.Type
End Property

Public Property Let DataType(ByVal dType As adStreamTypeEnum)
    Stream.Type = dType
End Property

Public Function EOS() As Boolean
    EOS = Stream.EOS
End Function

Public Property Get LineSeparator() As adLineSeparatorsEnum
    LineSeparator = Stream.LineSeparator
End Property

Public Property Let LineSeparator(ByVal SeparatorType As adLineSeparatorsEnum)
    Stream.LineSeparator = SeparatorType
End Property

Public Property Get Mode() As adConnectModeEnum
    Mode = Stream.Mode
End Property

Public Property Let Mode(ByVal Mode As adConnectModeEnum)
    Stream.Mode = Mode
End Property

Public Property Get Position() As Long
    Position = Stream.Position
End Property

Public Property Let Position(ByVal pos As Long)
    Stream.Position = pos
End Property

Public Property Get Size() As Long
    Size = Stream.Size
End Property

Private Sub Class_Initialize()
    Set Stream = CreateObject("ADODB.Stream")
End Sub

Private Sub Class_Terminate()
    If Stream.State <> adStateEnum.adStateClosed Then
        Stream.Close
    End If
End Sub

Public Function Cancel() As Boolean
    On Error GoTo Catch
    Stream.Cancel
    Cancel = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.Cancel:" & err.Description
    Cancel = False
End Function

Public Function CloseStream() As Boolean
    On Error GoTo Catch
    Stream.Close
    CloseStream = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.CloseStream:" & err.Description
    CloseStream = False
End Function

Public Function CopyTo(ByRef DestStream As Object, Optional NumChars As Integer = -1)
    On Error GoTo Catch
    Stream.CopyTo DestStream, NumChars
    CopyTo = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.CopyTo:" & err.Description
    CopyTo = False
End Function

Public Function Flush() As Boolean
    On Error GoTo Catch
    Stream.Flush
    Flush = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.Flush:" & err.Description
    Flush = False
End Function

Public Function LoadFromFile(ByVal FileName As String) As Boolean
    On Error GoTo Catch
    Stream.LoadFromFile FileName
    LoadFromFile = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.LoadFromFile:" & err.Description
    LoadFromFile = False
End Function

Public Function OpenStream( _
       Optional ByVal Source As Variant, _
       Optional ByVal Mode As adConnectModeEnum = adConnectModeEnum.adModeUnknown, _
       Optional ByVal OpenOptions As adStreamOpenOptionsEnum = adStreamOpenOptionsEnum.adOpenStreamUnspecified, _
       Optional ByVal UserName As String, _
       Optional ByVal Password As String) As Boolean
    On Error GoTo Catch
    Stream.Open Source, Mode, OpenOptions, UserName, Password
    OpenStream = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.OpenStream:" & err.Description
    OpenStream = False
End Function

Public Function ReadBinary(Optional NumBytes As Long = -1) As Variant
    ReadBinary = Stream.Read(NumBytes)
End Function

Public Function ReadText(Optional ByVal NumChars As Long = -1) As String
    ReadText = Stream.ReadText(NumChars)
End Function

Public Function SaveToFile( _
       ByVal FileName As String, _
       Optional ByVal SaveOptions As adSaveOptionsEnum = adSaveOptionsEnum.adSaveCreateNotExist _
       ) As Boolean
    On Error GoTo Catch
    Stream.SaveToFile FileName, SaveOptions
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
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.SaveToFile:" & err.Description
    SaveToFile = False
End Function

Public Function SetCharset(ByVal char As adCharSetEnum)
    Select Case char
    Case adCharSetEnum.ASCII
        Stream.Charset = "ascii"
    Case adCharSetEnum.ISO_8859_1
        Stream.Charset = "iso-8859-1"
    Case adCharSetEnum.Shift_Jis
        Stream.Charset = "shift_jis"
    Case adCharSetEnum.UTF_7
        Stream.Charset = "utf-7"
    Case adCharSetEnum.UTF_8, adCharSetEnum.UTF_8n
        Stream.Charset = "utf-8"
    Case adCharSetEnum.UTF_16LE
        Stream.Charset = "unicode"
    Case adCharSetEnum.UTF_16BE
        Stream.Charset = "unicodeFFFE"
    End Select
    If char = UTF_8n Then
        utf8NoBOM = True
    Else
        utf8NoBOM = False
    End If
End Function

Function SetEOS() As Boolean
    On Error GoTo Catch
    Stream.SetEOS
    SetEOS = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.SetEOS:" & err.Description
    SetEOS = False
End Function

Public Function SkipLine() As Boolean
    On Error GoTo Catch
    Stream.SkipLine
    SkipLine = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.SkipLine:" & err.Description
    SkipLine = False
End Function

Public Function WriteBinary(ByVal Buffer As Variant) As Boolean
    On Error GoTo Catch
    Stream.Write Buffer
    WriteBinary = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.WriteBinary:" & err.Description
    WriteBinary = False
End Function

Public Function WriteText( _
       ByVal Data As String, _
       Optional Options As adStreamWriteEnum = adStreamWriteEnum.adWriteChar _
       ) As Boolean
    On Error GoTo Catch
    Stream.WriteText Data, Options
    WriteText = True
    Exit Function
Catch:
    If Stream_DEBUG Then Debug.Print "[" & Format$(Date, "yyyy/mm/dd hh:mm:ss") & "] Stream.WriteText:" & err.Description
    WriteText = False
End Function


