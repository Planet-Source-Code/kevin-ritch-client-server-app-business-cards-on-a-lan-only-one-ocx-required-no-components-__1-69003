Attribute VB_Name = "V8ClientModule"
Global AutoListing$
Global CreatedNewBusinessCard As Boolean
Global EditingBusinessCard As Boolean
Global ContactData
Global CurrentContactRecNum As String
Global tb As String
Global SnStr As String
Global HTTPPath$

Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long
Global Const SWP_NOMOVE = 2
Global Const SWP_NOSIZE = 1
Global Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2

Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal sURL As String, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Public Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
    "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer _
    As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal _
    lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Const IF_FROM_CACHE = &H1000000
Public Const IF_MAKE_PERSISTENT = &H2000000
Public Const IF_NO_CACHE_WRITE = &H4000000
       
Public Const FLD_MrMrsMs = 1
Public Const FLD_FName = 2
Public Const FLD_LName = 3
Public Const FLD_Company = 4
Public Const FLD_Phone = 5
Public Const FLD_Extension = 6
Public Const FLD_Fax = 7
Public Const FLD_Mobile = 8
Public Const FLD_Home = 9
Public Const FLD_EMail = 10
Public Const FLD_Website = 11
Public Const FLD_Addr1 = 12
Public Const FLD_Addr2 = 13
Public Const FLD_City = 14
Public Const FLD_State = 15
Public Const FLD_Zip = 16
Public Const FLD_IDStatus = 17
Public Const FLD_Owner = 18
Public Const FLD_DirectLine = 19
Public Const FLD_Department = 20
Public Const FLD_JobTitle = 21
       
       
       
Private Const BUFFER_LEN = 1024 ' 1 ' 128

Public Function GetUrlSource(sURL As String) As String
    Screen.MousePointer = 11
    Dim sBuffer As String * BUFFER_LEN, iResult As Integer, sData As String
    Dim hInternet As Long, hSession As Long, lReturn As Long

    'get the handle of the current internet connection
    hSession = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    'get the handle of the url
    If hSession Then hInternet = InternetOpenUrl(hSession, sURL, vbNullString, 0, IF_NO_CACHE_WRITE, 0)
    'if we have the handle, then start reading the web page
    If hInternet Then
    
        'get the first chunk & buffer it.
        iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
        sData = sBuffer
        'if there's more data then keep reading it into the buffer
        Do While lReturn <> 0
            iResult = InternetReadFile(hInternet, sBuffer, BUFFER_LEN, lReturn)
            sData = sData + Mid(sBuffer, 1, lReturn)
        Loop
    End If
   
    'close the URL
    iResult = InternetCloseHandle(hInternet)

    GetUrlSource = sData
    Screen.MousePointer = Default
End Function

Sub HardDiskSerial()
 Dim volname As String   ' receives volume name of C:
 Dim sn As Long          ' receives serial number of C:
 Dim maxcomplen As Long  ' receives maximum component length
 Dim sysflags As Long    ' receives file system flags
 Dim sysname As String   ' receives the file system name
 Dim retval As Long      ' return value
 volname = Space(256)
 sysname = Space(256)
 retval = GetVolumeInformation("C:\", volname, Len(volname), sn, maxcomplen, sysflags, sysname, Len(sysname))
 SnStr = Trim(Hex(sn))
 SnStr = String(8 - Len(SnStr), "0") & SnStr
 SnStr = Left(SnStr, 4) & "-" & Right(SnStr, 4)
End Sub
