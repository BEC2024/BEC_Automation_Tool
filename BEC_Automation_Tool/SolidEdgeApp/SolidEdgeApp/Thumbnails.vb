Imports System.IO
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Text
Imports SeThumbnailLib

Friend Class Thumbnails



    Private Shared ReadOnly IID_ISHELLFOLDER As New Guid("000214E6-0000-0000-C000-000000000046")
    Private Shared ReadOnly IID_IEXTRACTIMAGE As New Guid("BB2E617C-0920-11d1-9A0B-00C04FC2D6C1")

    Public Shared Function ExtractThumbNail(ByVal file As FileInfo) As Bitmap
        Return ExtractThumbNail(file, New System.Drawing.Size(100, 100))
    End Function

    Public Shared Function ExtractThumbNail(ByVal file As FileInfo, ByVal imageSize As Size) As Bitmap

        Dim thumbnail As Bitmap = Nothing
        'Dim alloc As IMalloc = Nothing
        Dim folder As IShellFolder = Nothing
        Dim item As IShellFolder = Nothing
        Dim pidlFolder As IntPtr = IntPtr.Zero
        Dim hBmp As IntPtr = IntPtr.Zero
        Dim extractImage As IExtractImage = Nothing
        Dim pidl As IntPtr = IntPtr.Zero

        If (file.Exists) Then

            Try
                SHGetDesktopFolder(folder)

                If Not folder Is Nothing Then

                    Dim cParsed As Integer = 0
                    Dim pdwAttrib As Integer = 0

                    Dim HR As Integer = folder.ParseDisplayName(IntPtr.Zero, IntPtr.Zero,
                     file.Directory.FullName, cParsed, pidlFolder,
                     pdwAttrib)
                    If HR < S_OK Then Return Nothing

                    If Not pidlFolder.Equals(IntPtr.Zero) Then

                        HR = folder.BindToObject(pidlFolder, IntPtr.Zero,
                            IID_ISHELLFOLDER, item)
                        If HR < S_OK Then Return Nothing

                        If Not item Is Nothing Then

                            HR = item.ParseDisplayName(IntPtr.Zero, IntPtr.Zero,
                                    file.Name, 0, pidl, 0)
                            Marshal.ThrowExceptionForHR(HR)

                            Dim prgf As Integer = 0
                            HR = item.GetUIObjectOf(0, 1, New IntPtr() {pidl},
                                IID_IEXTRACTIMAGE, prgf, extractImage)
                            If HR < S_OK Then Return Nothing

                            If Not extractImage Is Nothing Then
                                Dim location As New StringBuilder(MAX_PATH, MAX_PATH)

                                Dim priority As Integer = 0
                                Dim requestedColorDepth As Integer = 32

                                Dim uFlags As Integer = IEIFLAG.IEIFLAG_ASPECT Or
                                    IEIFLAG.IEIFLAG_ORIGSIZE Or IEIFLAG.IEIFLAG_QUALITY

                                HR = extractImage.GetLocation(location, location.Capacity,
                                        priority, imageSize, requestedColorDepth,
                                        uFlags)
                                If HR < S_OK Then Return Nothing

                                HR = extractImage.Extract(hBmp)
                                If HR < S_OK Then Return Nothing
                                If Not hBmp.Equals(IntPtr.Zero) Then
                                    thumbnail = Bitmap.FromHbitmap(hBmp)
                                End If
                            End If
                        End If
                    End If
                End If
            Finally

                If Not hBmp.Equals(IntPtr.Zero) Then DeleteObject(hBmp)
                If Not pidlFolder.Equals(IntPtr.Zero) Then
                    Marshal.FreeCoTaskMem(pidlFolder)
                End If
                If Not extractImage Is Nothing Then
                    Marshal.ReleaseComObject(extractImage)
                    extractImage = Nothing
                End If
                If Not item Is Nothing Then
                    Marshal.ReleaseComObject(item)
                    item = Nothing
                End If
                If Not folder Is Nothing Then
                    Marshal.ReleaseComObject(folder)
                    folder = Nothing
                End If
            End Try
        End If
        Return thumbnail
    End Function

    Private Const S_OK As Integer = 0
    Public Shared ReadOnly IID_ContextMenu As New Guid("000214e4-0000-0000-c000-000000000046")

    Private Const MAX_PATH As Integer = 260

    <DllImport("gdi32", CharSet:=CharSet.Auto)>
    Private Shared Function DeleteObject(ByVal hObject As IntPtr) As Integer
    End Function

    Private Declare Auto Function SHGetDesktopFolder Lib "shell32" (
            ByRef ppshf As IShellFolder) As Integer

End Class

<ComImportAttribute(),
GuidAttribute("BB2E617C-0920-11d1-9A0B-00C04FC2D6C1"),
InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)>
Friend Interface IExtractImage

    <PreserveSig()>
    Function GetLocation(<Out(), MarshalAs(UnmanagedType.LPWStr)> ByVal pszPathBuffer As StringBuilder,
     ByVal cch As Integer, ByRef pdwPriority As Integer, ByRef prgSize As Size,
     ByVal dwRecClrDepth As Integer, ByRef pdwFlags As Integer) As Integer

    <PreserveSig()>
    Function Extract(<Out()> ByRef phBmpThumbnail As IntPtr) As Integer

End Interface

<ComImportAttribute(),
 InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
 Guid("000214E6-0000-0000-C000-000000000046")>
Friend Interface IShellFolder
    <PreserveSig()>
    Function ParseDisplayName(
        ByVal hwndOwner As IntPtr,
        ByVal pbcReserved As IntPtr,
        <MarshalAs(UnmanagedType.LPWStr)>
        ByVal lpszDisplayName As String,
        ByRef pchEaten As Integer,
        ByRef ppidl As IntPtr,
        ByRef pdwAttributes As Integer) As Integer

    <PreserveSig()>
    Function EnumObjects(
        ByVal hwndOwner As Integer,
        <MarshalAs(UnmanagedType.U4)> ByVal _
        grfFlags As Integer,
        ByRef ppenumIDList As IntPtr) As Integer

    <PreserveSig()>
    Function BindToObject(
        ByVal pidl As IntPtr,
        ByVal pbcReserved As IntPtr,
        ByRef riid As Guid,
        ByRef ppvOut As IShellFolder) As Integer
    'IShellFolder) As Integer

    <PreserveSig()>
    Function BindToStorage(
        ByVal pidl As IntPtr,
        ByVal pbcReserved As IntPtr,
        ByRef riid As Guid,
        ByVal ppvObj As IntPtr) As Integer

    <PreserveSig()>
    Function CompareIDs(
        ByVal lParam As IntPtr,
        ByVal pidl1 As IntPtr,
        ByVal pidl2 As IntPtr) As Integer

    <PreserveSig()>
    Function CreateViewObject(
        ByVal hwndOwner As IntPtr,
        ByRef riid As Guid,
        ByRef ppvOut As IntPtr) As Integer
    'IUnknown) As Integer

    <PreserveSig()>
    Function GetAttributesOf(
        ByVal cidl As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)>
        ByVal apidl() As IntPtr,
        ByRef rgfInOut As Integer) As Integer

    <PreserveSig()>
    Function GetUIObjectOf(
        ByVal hwndOwner As Integer,
        ByVal cidl As Integer,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)>
        ByVal apidl() As IntPtr,
        ByRef riid As Guid,
        <Out()> ByRef prgfInOut As Integer,
        <Out(), MarshalAs(UnmanagedType.IUnknown)> ByRef ppvOut As Object) As Integer
    'ByRef ppvOut As IUnknown) As Integer
    'ByRef ppvOut As IDropTarget) As Integer

    <PreserveSig()>
    Function GetDisplayNameOf(
        ByVal pidl As IntPtr,
        <MarshalAs(UnmanagedType.U4)>
        ByVal uFlags As Integer,
        ByVal lpName As IntPtr) As Integer

    <PreserveSig()>
    Function SetNameOf(
        ByVal hwndOwner As Integer,
        ByVal pidl As IntPtr,
        <MarshalAs(UnmanagedType.LPWStr)> ByVal _
        lpszName As String,
        <MarshalAs(UnmanagedType.U4)> ByVal _
        uFlags As Integer,
        ByRef ppidlOut As IntPtr) As Integer

End Interface

<Flags()>
Friend Enum IEIFLAG
    IEIFLAG_ASYNC = &H1     ' ask the extractor if it supports ASYNC extract (free threaded)
    IEIFLAG_CACHE = &H2      'returned from the extractor if it does NOT cache the thumbnail
    IEIFLAG_ASPECT = &H4      ' passed to the extractor to beg it to render to the aspect ratio of the supplied rect
    IEIFLAG_OFFLINE = &H8     ' if the extractor shouldn't hit the net to get any content needed for the rendering
    IEIFLAG_GLEAM = &H10     'does the image have a gleam ? this will be returned if it does
    IEIFLAG_SCREEN = &H20      ' render as if for the screen  (this is exlusive with IEIFLAG_ASPECT )
    IEIFLAG_ORIGSIZE = &H40      ' render to the approx size passed, but crop if neccessary
    IEIFLAG_NOSTAMP = &H80      ' returned from the extractor if it does NOT want an icon stamp on the thumbnail
    IEIFLAG_NOBORDER = &H100      'returned from the extractor if it does NOT want an a border around the thumbnail
    IEIFLAG_QUALITY = &H200      ' passed to the Extract method to indicate that a slower, higher quality image is desired, re-compute the thumbnail
End Enum
