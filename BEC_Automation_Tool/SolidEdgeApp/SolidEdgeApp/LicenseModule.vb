Imports System.Management



Module LicenseModule

    Public privateRegistrationKey As String = "a*U7$hY5"
    Public privateLicenseKey As String = "j&G4Hy6$"

    Public Function getMacId() As String
        Dim mc As New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        Dim MACAddress As String = [String].Empty
        For Each mo As ManagementObject In moc
            ' only return MAC Address from first card
            If CBool(mo("IPEnabled")) = True Then
                MACAddress = mo("MacAddress").ToString()
                Exit For
            End If
            mo.Dispose()
        Next

        MACAddress = MACAddress.Replace(":", "")
        Return MACAddress
    End Function

    Public Function GetMotherBoardID() As String
        Dim strMotherBoardID As String = String.Empty
        Dim query As New SelectQuery("Win32_BaseBoard")
        Dim search As New ManagementObjectSearcher(query)

        For Each info As ManagementObject In search.Get()
            strMotherBoardID = info("SerialNumber").ToString()
        Next

        Return strMotherBoardID
    End Function

    Public Function Encrypt(ByRef pPassPhrase As String, ByVal pTextToEncrypt As String) As String
        If pPassPhrase.Length > 16 Then
            'limitation of the encryption mechanism
            pPassPhrase = pPassPhrase.Substring(0, 16)
        End If

        If pTextToEncrypt.Trim.Length = 0 Then
            'the Text to encrypt not set!!!
            Return String.Empty
        End If

        Dim skey As New Encryption.Data(pPassPhrase)
        Dim sym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
        Dim objEncryptedData As Encryption.Data
        objEncryptedData = sym.Encrypt(New Encryption.Data(pTextToEncrypt), skey)
        Return objEncryptedData.ToHex
    End Function


    Public Function Decrypt(ByRef pPassPhrase As String, ByVal pHexStream As String) As String
        Try
            Dim objSym As New Encryption.Symmetric(Encryption.Symmetric.Provider.Rijndael)
            Dim encryptedData As New Encryption.Data
            encryptedData.Hex = pHexStream
            Dim decryptedData As Encryption.Data
            decryptedData = objSym.Decrypt(encryptedData, New Encryption.Data(pPassPhrase))
            Return decryptedData.Text
        Catch
            Return Nothing
        End Try
    End Function
End Module
