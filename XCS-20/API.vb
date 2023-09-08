Module API
    Public Structure SECURITY_ATTRIBUTES
        Dim nLength As Long
        Dim lpSecurityDescriptor As Long
        Dim bInheritHandle As Long
    End Structure
    Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
End Module
