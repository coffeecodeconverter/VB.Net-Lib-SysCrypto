# VB.Net-Lib-SysCrypto
VB.Net Module for System, Network, and User Info, plus Encryption, and Encoding functions - can be used for a software licensing system

# 🔐 VB.NET System Info, Encryption & Licensing Module

A powerful and versatile VB.NET module for gathering system, user, network, and hardware details, while offering strong cryptographic tools such as AES/RSA encryption, hashing, UUID generation, and OTP (TOTP/HOTP) support.

Originally created for a flexible **software licensing system**, but useful across many domains including telemetry, authentication, diagnostics, and hardware binding.

---
# 🔐 Sample Application
The Project comes with an .exe that demonstrates using all of the functions for various tasks. 
Making it easy to learn how to integrate this module into your existing projects. 
It includes demonstrations on generating licenses, either as .LIC files or using serial keys. 

---

## ✅ Requirements

- **.NET Framework:** 4.6.1 or higher
- **NuGet Packages:**
  - `Newtonsoft.Json`
  - `BouncyCastle`
  - `OTP.Net`

---

## ✨ Features

### 🔐 AES Encryption / Decryption

```vb.net
' String encryption/decryption
Await EncryptAsync(data, EncryptionKey, EncryptionIV)
Await DecryptAsync(data, EncryptionKey, EncryptionIV)

' File encryption/decryption
Await EncryptFileAsync(inputFilePath, outputFilePath, EncryptionKey, EncryptionIV)
Await DecryptFileAsync(inputFilePath, outputFilePath, EncryptionKey, EncryptionIV)
```

### 🔑 RSA Cryptography

```vb.net
' Generate and manage key pairs
Await GenerateRsaKeysAsync(keySize, seed)
RSA_ExportGeneratedKeyPairs()
RSA_ImportGeneratedKeyPairs()

' Encrypt/decrypt
RsaEncrypt(plainText, publicKeyPem, salt)
RsaDecrypt(cipherText, privateKeyPem, salt)

' Sign/verify
RsaSignData(data, privateKeyPem, salt)
RsaVerifySignature(data, signature, publicKeyPem, salt)

' Key inspection
RSA_ExtractLastCharsFromKey(publicOrPrivateKey, numOfChars)
```

### 🧠 System Information

```vb.net
MsgBox(Await GetThe_Computer("Hostname"))
   MsgBox(Await GetThe_Computer("Domain"))
   MsgBox(Await GetThe_Computer("Make"))
   MsgBox(Await GetThe_Computer("Model"))
   MsgBox(Await GetThe_Computer("SystemModel"))
   MsgBox(Await GetThe_Computer("SystemModelFull"))
   MsgBox(Await GetThe_Computer("BIOSSerial"))
   MsgBox(Await GetThe_Computer("MotherboardSerial"))
   MsgBox(Await GetThe_Computer("MotherBoardBaseBoardProduct"))
   MsgBox(Await GetThe_Computer("UUID"))
   MsgBox(Await GetThe_Computer("CPUName"))
   MsgBox(Await GetThe_Computer("CPUCount"))
   MsgBox(Await GetThe_Computer("CPUCores"))
   MsgBox(Await GetThe_Computer("CPUThreads"))
   MsgBox(Await GetThe_Computer("CPUSpeed"))
   MsgBox(Await GetThe_Computer("CPUArchitecture"))
   MsgBox(Await GetThe_Computer("RAMSize"))
   MsgBox(Await GetThe_Computer("RAMSpeed"))
   MsgBox(Await GetThe_Computer("OSName"))
   MsgBox(Await GetThe_Computer("OSBuild"))
   MsgBox(Await GetThe_Computer("OSVersion"))
   MsgBox(Await GetThe_Computer("OSArchitecture"))
   MsgBox(Await GetThe_Computer("OSServiceChannel"))

   MsgBox(Await GetThe_HardDrive("Count"))
   MsgBox(Await GetThe_HardDrive("PNPDeviceID"))
   MsgBox(Await GetThe_HardDrive("DriveSerialNumber"))
   MsgBox(Await GetThe_HardDrive("VolumeSerialNumber"))
   MsgBox(Await GetThe_HardDrive("Model"))
   MsgBox(Await GetThe_HardDrive("DiskDeviceID"))
   MsgBox(Await GetThe_HardDrive("SizeB"))
   MsgBox(Await GetThe_HardDrive("SizeGB"))
   MsgBox(Await GetThe_HardDrive("FreeSpaceB"))
   MsgBox(Await GetThe_HardDrive("FreeSpaceGB"))
   MsgBox(Await GetThe_HardDrive("DriveLetter"))
   MsgBox(Await GetThe_HardDrive("VolumeName"))
   MsgBox(Await GetThe_HardDrive("FileSystem"))

   MsgBox(Await GetThe_Network("NICDescription"))
   MsgBox(Await GetThe_Network("IPv4Address"))
   MsgBox(Await GetThe_Network("MACAddress"))
   MsgBox(Await GetThe_Network("SubnetMask"))
   MsgBox(Await GetThe_Network("DefaultGateway"))
   MsgBox(Await GetThe_Network("DNSDomainSuffix"))
   Or get all network adapter details in a dictionary
   Dim dictAllNetworkAdapterDetails As Dictionary(Of String, Dictionary(Of String, String)) = Await GetDictionary_NetworkDetails()

   MsgBox(Await GetThe_User("Username"))
   MsgBox(Await GetThe_User("DomainAndUsername"))
   MsgBox(Await GetThe_User("AdminRights"))
```

### 🧪 Encoding & Utilities
```vb.net
Base64String_Encode(passedString)
Base64String_Decode(passedString)

Base32String_Encode(byteData)
Base32String_Decode(passedString)

SplitAndReverseString(passedString)

GenerateUUID()
GenerateRandomString(length)
GetHash(passedString)
GetFileHash(filePath)

CopyStringToClipboard(text)

MsgBox(Await GetFingerprint_Computer())
MsgBox(Await GetFingerprint_HardDrives())
MsgBox(Await GetFingerprint_HardDriveAppRunsFrom())

GetSerials_HardDrivesAll()
GetAllDriveLettersAsync()

GenerateTOTP(base32Secret, Optional stepSeconds)
GenerateHOTP(base32Secret, counter)
```


### Full Module 
```vb.net
Imports System.Security.Cryptography
Imports System.Text
Imports System.Management                           ' need to install this package via nuget
Imports System.IO
Imports System.Threading.Tasks
Imports System.Security.Principal
Imports Newtonsoft.Json
Imports System.Net
Imports Microsoft.Win32
Imports System.Net.NetworkInformation               ' install package via nuget

Imports Org.BouncyCastle.Crypto                     ' install BouncyCastle package via nuget
Imports Org.BouncyCastle.Crypto.Digests
Imports Org.BouncyCastle.Crypto.Generators
Imports Org.BouncyCastle.Security
Imports Org.BouncyCastle.OpenSsl
Imports Org.BouncyCastle.Crypto.Engines
Imports Org.BouncyCastle.Crypto.Prng
Imports Newtonsoft.Json.Linq
Imports System.Runtime.CompilerServices

Imports OtpNet                                      ' install via nuget package "OTP.net"



' This class represents an RSA key pair and will be used for JSON serialization
Public Class RSAKeyPair
    Public Property KeyIdentifier As Integer
    Public Property PublicKey As String
    Public Property PrivateKey As String
    Public Property Description As String
    Public Property Seed As String
    Public Property Salt As String
    Public Property KeySize As Integer
End Class



Module LicenseGenerator


    ' Global Variables
    Public moduleVersion_LicenseGen As String = "1.0.0.25"
    Public dictComputerDetails As Dictionary(Of String, String) = Nothing
    Public RSAPublicKey As String = ""
    Public RSAPrivateKey As String = ""
    Public RSAKeyPairs As New Dictionary(Of Integer, Dictionary(Of String, String))()




    ' FUNCTIONS TO USE
    ' ----------------------------------
    Public Async Function GetThe_Computer(ByVal detail As String) As Task(Of String)
        ' Initialize a global variable, as this information never changes whilst machine is running, once known. 
        If dictComputerDetails Is Nothing Then
            dictComputerDetails = Await GetDictionary_ComputerDetails()
        End If

        If dictComputerDetails.ContainsKey(detail) Then
            Return dictComputerDetails(detail)
        Else
            Debug.WriteLine("Error: Specified detail not found in dictionary. Accepted values are:" & vbCrLf &
                "- Hostname" & vbCrLf &
                "- Domain" & vbCrLf &
                "- Make" & vbCrLf &
                "- Model" & vbCrLf &
                "- SystemModel" & vbCrLf &
                "- SystemModelFull" & vbCrLf &
                "- BIOSSerial" & vbCrLf &
                "- MotherboardSerial" & vbCrLf &
                "- MotherBoardBaseBoardProduct" & vbCrLf &
                "- UUID" & vbCrLf &
                "- CPUCount" & vbCrLf &
                "- CPUName" & vbCrLf &
                "- CPUCores" & vbCrLf &
                "- CPUThreads" & vbCrLf &
                "- CPUSpeed" & vbCrLf &
                "- CPUArchitecture" & vbCrLf &
                "- RAMSize" & vbCrLf &
                "- RAMSpeed" & vbCrLf &
                "- OSName" & vbCrLf &
                "- OSVersion" & vbCrLf &
                "- OSBuild" & vbCrLf &
                "- OSArchitecture" & vbCrLf &
                "- OSServiceChannel")
            dictComputerDetails = Nothing
            Return Nothing
        End If

    End Function



    Public Async Function GetThe_HardDrive(ByVal detail As String, Optional ByVal specificDriveLetter As String = "") As Task(Of String)

        Dim driveLetter As String = ""
        If specificDriveLetter = "" Then
            ' Get the path where the application is running from
            Dim appPath As String = Application.StartupPath
            driveLetter = Path.GetPathRoot(appPath).TrimEnd("\"c)
        Else
            ' this way, doesnt matter is user passes "C", "C:", or "C:\", or even "C:\Anypath"
            driveLetter = specificDriveLetter.Substring(0, 1) & ":"
        End If

        Dim dictDriveDetails As Dictionary(Of String, String) = Await GetDictionary_HardDriveDetails(driveLetter)

        Try
            If dictDriveDetails.ContainsKey(detail) Then
                Return dictDriveDetails(detail)
            Else
                Debug.WriteLine("Error: Specified detail not found in dictionary. Accepted values are:" & vbCrLf &
                    "- Count" & vbCrLf &
                    "- PNPDeviceID" & vbCrLf &
                    "- DriveSerialNumber" & vbCrLf &
                    "- VolumeSerialNumber" & vbCrLf &
                    "- Model" & vbCrLf &
                    "- DiskDeviceID" & vbCrLf &
                    "- SizeB" & vbCrLf &
                    "- SizeGB" & vbCrLf &
                    "- FreeSpaceB" & vbCrLf &
                    "- FreeSpaceGB" & vbCrLf &
                    "- DriveLetter" & vbCrLf &
                    "- VolumeName" & vbCrLf &
                    "- FileSystem")
                dictDriveDetails = Nothing
                Return Nothing
            End If
        Catch ex As Exception
            MessageBox.Show("Error checking hard drive information:" & vbCrLf & "Detail: " & detail & vbCrLf & vbCrLf & "Message:" & vbCrLf & ex.Message)
        End Try

    End Function






    Public Async Function GetThe_Network(Optional ByVal requestedDetail As String = "IPv4Address") As Task(Of String)

        ' Update global variable each time as network details can change
        Dim dictNetworkDetails As Dictionary(Of String, Dictionary(Of String, String)) = Await GetDictionary_NetworkDetails()

        '''''' Debugging: Print all the adapters in the debug output for inspection
        '''''Debug.WriteLine("All Network Adapters:")
        '''''For Each adapter As KeyValuePair(Of String, Dictionary(Of String, String)) In dictNetworkDetails
        '''''    Debug.WriteLine($"{vbCrLf}Adapter Name: {adapter.Key}")
        '''''    For Each detail As KeyValuePair(Of String, String) In adapter.Value
        '''''        Debug.WriteLine($"  {detail.Key}: {detail.Value}")
        '''''    Next
        '''''Next
        '''''Debug.WriteLine("End of Network Adapters")

        Dim result As String = Nothing

        For Each adapter As KeyValuePair(Of String, Dictionary(Of String, String)) In dictNetworkDetails
            ' This excludes virtual, bluetooth, and Loopback addresses 
            If Not IsValidAdapter(adapter.Key, adapter.Value("NICDescription")) Then
                Continue For
            End If

            If adapter.Value("IPv4Address").Substring(0, 3) = "169" Then
                Continue For
            End If

            If adapter.Value.ContainsKey(requestedDetail) Then
                'result = $"{adapter.Key}: {adapter.Value(requestedDetail)}"
                result = adapter.Value(requestedDetail)
                Exit For
            End If
        Next

        ' If the requested detail isn't found in any adapter
        If result Is Nothing Then
            Debug.WriteLine($"Error: Specified detail '{requestedDetail}' not found in dictionary. Accepted values are:" & vbCrLf &
                "- NICDescription" & vbCrLf &
                "- IPv4Address" & vbCrLf &
                "- MACAddress" & vbCrLf &
                "- SubnetMask" & vbCrLf &
                "- DefaultGateway" & vbCrLf &
                "- DNSDomainSuffix")
            Return Nothing
        End If

        Return result
    End Function








    Public Async Function GetThe_User(Optional ByVal detail As String = "username") As Task(Of String)
        Try
            Return Await Task.Run(Function()
                                      Dim currentUser As WindowsIdentity = WindowsIdentity.GetCurrent()
                                      Dim username As String = currentUser.Name.Split("\"c).Last()
                                      Dim domainAndUsername As String = currentUser.Name
                                      Dim adminRights As Boolean = False
                                      Dim principal As New WindowsPrincipal(currentUser)
                                      If principal.IsInRole(WindowsBuiltInRole.Administrator) Then
                                          adminRights = True
                                      End If

                                      Select Case detail.ToLower()
                                          Case "username"
                                              Return username
                                          Case "domainandusername"
                                              Return domainAndUsername
                                          Case "adminrights"
                                              Return If(adminRights, "Yes", "No")
                                          Case Else
                                              Return "Invalid detail specified"
                                      End Select
                                  End Function)
        Catch ex As Exception
            Debug.WriteLine("Error: " & ex.Message)
            Return Nothing
        End Try
    End Function










    Public Async Function GetFingerprint_HardDrives() As Task(Of List(Of String))
        Dim serials As List(Of String) = Await GetSerials_HardDrivesAll()
        Dim hashes As New List(Of String)()
        For Each serial As String In serials
            Dim hashString As String = GetHash(serial)
            hashes.Add(hashString)
            Debug.WriteLine(serial & "  Hash:" & hashString)
        Next
        Return hashes
    End Function






    Public Async Function GetFingerprint_HardDriveAppRunsFrom() As Task(Of String)
        Dim driveLetterAppRunsFrom As String = Application.StartupPath.Substring(0, 1) & ":"
        Dim volumeSerialAppRunsFrom As String = Await GetThe_HardDrive("DriveSerialNumber", driveLetterAppRunsFrom)
        Debug.WriteLine("volumeSerialAppRunsFrom: " & volumeSerialAppRunsFrom)
        Dim hashString As String = GetHash(volumeSerialAppRunsFrom)
        Return hashString
    End Function







    Public Async Function GetFingerprint_Computer(Optional includeMake As Boolean = True,
                                                  Optional includeModel As Boolean = True,
                                                  Optional includeSystemModel As Boolean = True,
                                                  Optional includeSystemModelFull As Boolean = True,
                                                  Optional includeBIOSSerial As Boolean = True,
                                                  Optional includeMotherboardSerial As Boolean = True,
                                                  Optional includeMotherboardBaseBoardProduct As Boolean = True,
                                                  Optional includeMachineUUID As Boolean = True) As Task(Of String)

        Dim concatenatedInfo As New StringBuilder()

        If includeMake = True Then
            concatenatedInfo.Append("Make=" & Await GetThe_Computer("Make") & ";")
        End If

        If includeModel = True Then
            concatenatedInfo.Append("Model=" & Await GetThe_Computer("Model") & ";")
        End If

        If includeSystemModel = True Then
            concatenatedInfo.Append("SystemModel=" & Await GetThe_Computer("SystemModel") & ";")
        End If

        If includeSystemModelFull = True Then
            concatenatedInfo.Append("SystemModelFull=" & Await GetThe_Computer("SystemModelFull") & ";")
        End If

        If includeBIOSSerial = True Then
            concatenatedInfo.Append("BIOSSerial=" & Await GetThe_Computer("BIOSSerial") & ";")
        End If

        If includeMotherboardSerial = True Then
            concatenatedInfo.Append("MotherboardSerial=" & Await GetThe_Computer("MotherboardSerial") & ";")
        End If

        If includeMotherboardBaseBoardProduct = True Then
            concatenatedInfo.Append("MotherBoardBaseBoardProduct=" & Await GetThe_Computer("MotherBoardBaseBoardProduct") & ";")
        End If

        If includeMachineUUID = True Then
            concatenatedInfo.Append("UUID=" & Await GetThe_Computer("UUID") & ";")
        End If

        Dim hashString As String = GetHash(concatenatedInfo.ToString())

        Return hashString
    End Function








    Public Function GenerateRandomString(length As Integer, Optional characterList As String = Nothing) As String
        Dim characters As String
        If characterList Is Nothing Then
            characters = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"   ' valid characters
        Else
            characters = characterList
        End If
        Dim random As New Random()
        Dim result As New Text.StringBuilder()

        For i As Integer = 1 To length
            ' Get a random index
            Dim index As Integer = random.Next(0, characters.Length)
            result.Append(characters(index))
        Next

        Return result.ToString()
    End Function




    Public Function MinifyJson(jsonString As String) As String
        Dim parsedJson As JObject = JObject.Parse(jsonString)
        Return JsonConvert.SerializeObject(parsedJson, Formatting.None)
    End Function






    ' Encryption 
    Public Function GetHash(passedString As String) As String
        Using sha256 As SHA256 = SHA256.Create()
            Dim hashBytes As Byte() = sha256.ComputeHash(Encoding.UTF8.GetBytes(passedString))
            Dim sb As New StringBuilder()
            For Each b As Byte In hashBytes
                sb.Append(b.ToString("x2")) ' Convert byte to hex
            Next
            Return sb.ToString()
        End Using
    End Function


    Public Function GetFileHash(filePath As String) As String
        If Not File.Exists(filePath) Then
            Throw New FileNotFoundException("The specified file does not exist.", filePath)
        End If
        Using sha256 As SHA256 = SHA256.Create()
            Using fileStream As FileStream = File.OpenRead(filePath)
                Dim hashBytes As Byte() = sha256.ComputeHash(fileStream)
                Dim sb As New StringBuilder()
                For Each b As Byte In hashBytes
                    sb.Append(b.ToString("x2")) ' Convert byte to hex format
                Next
                Return sb.ToString()
            End Using
        End Using
    End Function


    Private Function GetKeyFromPassphrase(passphrase As String) As Byte()
        Using sha256 As SHA256 = SHA256.Create()
            Return sha256.ComputeHash(Encoding.UTF8.GetBytes(passphrase))
        End Using
    End Function


    Private Function GetIvFromPassphrase(passphrase As String) As Byte()
        Using sha256 As SHA256 = SHA256.Create()
            Return sha256.ComputeHash(Encoding.UTF8.GetBytes(passphrase)).Take(16).ToArray()
        End Using
    End Function



    ' Encrypt
    Public Function EncryptAsync(plainText As String, privateKey As String, password As String) As Task(Of String)
        Return Task.Run(Function()
                            Try
                                Using aesAlg As Aes = Aes.Create()
                                    aesAlg.Key = GetKeyFromPassphrase(privateKey)
                                    aesAlg.IV = GetIvFromPassphrase(password)

                                    Using encryptor As ICryptoTransform = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV)
                                        Using msEncrypt As New MemoryStream()
                                            Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                                                Using swEncrypt As New StreamWriter(csEncrypt)
                                                    swEncrypt.Write(plainText)
                                                End Using
                                            End Using
                                            Return Convert.ToBase64String(msEncrypt.ToArray())
                                        End Using
                                    End Using
                                End Using
                            Catch ex As CryptographicException
                                Console.WriteLine("Cryptographic error: " & ex.StackTrace)
                                Return String.Empty
                            Catch ex As ArgumentNullException
                                Console.WriteLine("Argument null error: " & ex.StackTrace)
                                Return String.Empty
                            Catch ex As Exception
                                Console.WriteLine("An error occurred: " & ex.StackTrace)
                                Return String.Empty
                            End Try
                        End Function)
    End Function


    ' Decrypt
    Public Function DecryptAsync(encryptedText As String, privateKey As String, password As String) As Task(Of String)
        Return Task.Run(Function()
                            Try
                                Using aesAlg As Aes = Aes.Create()
                                    aesAlg.Key = GetKeyFromPassphrase(privateKey)
                                    aesAlg.IV = GetIvFromPassphrase(password)

                                    Using decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)
                                        Using msDecrypt As New MemoryStream(Convert.FromBase64String(encryptedText))
                                            Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                                                Using srDecrypt As New StreamReader(csDecrypt)
                                                    Return srDecrypt.ReadToEnd()
                                                End Using
                                            End Using
                                        End Using
                                    End Using
                                End Using
                            Catch ex As CryptographicException
                                Console.WriteLine("Cryptographic error - Invalid Private Key or Password: " & ex.StackTrace)
                                MsgBox("Cryptographic error - Invalid Private Key or Password: " & ex.StackTrace)
                                Return String.Empty
                            Catch ex As FormatException
                                Console.WriteLine("Format error: " & ex.StackTrace)
                                MsgBox("Format error: " & ex.StackTrace)
                                Return String.Empty
                            Catch ex As ArgumentNullException
                                Console.WriteLine("Argument null error: " & ex.StackTrace)
                                MsgBox("Argument null error: " & ex.StackTrace)
                                Return String.Empty
                            Catch ex As Exception
                                Console.WriteLine("An error occurred: " & ex.StackTrace)
                                MsgBox("An error occurred: " & ex.StackTrace)
                                Return String.Empty
                            End Try
                        End Function)
    End Function




    Public Function EncryptFileAsync(inputFilePath As String, outputFilePath As String, privateKey As String, password As String) As Task(Of Boolean)
        Return Task.Run(Function()
                            Try
                                Dim fileBytes As Byte() = File.ReadAllBytes(inputFilePath)

                                ' Encrypt the file bytes
                                Using aesAlg As Aes = Aes.Create()
                                    aesAlg.Key = GetKeyFromPassphrase(privateKey)
                                    aesAlg.IV = GetIvFromPassphrase(password)

                                    Using encryptor As ICryptoTransform = aesAlg.CreateEncryptor(aesAlg.Key, aesAlg.IV)
                                        Using msEncrypt As New MemoryStream()
                                            Using csEncrypt As New CryptoStream(msEncrypt, encryptor, CryptoStreamMode.Write)
                                                csEncrypt.Write(fileBytes, 0, fileBytes.Length)
                                            End Using

                                            ' Write the encrypted data to the output file
                                            File.WriteAllBytes(outputFilePath, msEncrypt.ToArray())
                                        End Using
                                    End Using
                                End Using

                                Return True
                            Catch ex As Exception
                                Debug.WriteLine("Error during file encryption: " & ex.StackTrace)
                                Return False
                            End Try
                        End Function)
    End Function




    Public Function DecryptFileAsync(inputFilePath As String, outputFilePath As String, privateKey As String, password As String) As Task(Of Boolean)
        Return Task.Run(Function()
                            Try
                                Dim encryptedFileBytes As Byte() = File.ReadAllBytes(inputFilePath)

                                ' Decrypt the file bytes
                                Using aesAlg As Aes = Aes.Create()
                                    aesAlg.Key = GetKeyFromPassphrase(privateKey)
                                    aesAlg.IV = GetIvFromPassphrase(password)

                                    Using decryptor As ICryptoTransform = aesAlg.CreateDecryptor(aesAlg.Key, aesAlg.IV)
                                        Using msDecrypt As New MemoryStream(encryptedFileBytes)
                                            Using csDecrypt As New CryptoStream(msDecrypt, decryptor, CryptoStreamMode.Read)
                                                Using msOutput As New MemoryStream()
                                                    csDecrypt.CopyTo(msOutput)

                                                    ' Write the decrypted data to the output file
                                                    File.WriteAllBytes(outputFilePath, msOutput.ToArray())
                                                End Using
                                            End Using
                                        End Using
                                    End Using
                                End Using

                                Return True
                            Catch ex As Exception
                                Debug.WriteLine("Error during file decryption: " & ex.StackTrace)
                                Return False
                            End Try
                        End Function)
    End Function








    ' Encryption - RSA

    Public Async Function GenerateRsaKeysAsync(Optional keySize As Integer = 1024, Optional seed As String = "") As Task(Of Boolean)
        Return Await Task.Run(Function()
                                  Try
                                      Dim random As SecureRandom

                                      If Not String.IsNullOrEmpty(seed) Then
                                          Dim seedBytes As Byte() = Encoding.UTF8.GetBytes(seed)
                                          Dim digestGenerator As New DigestRandomGenerator(New Sha256Digest())
                                          digestGenerator.AddSeedMaterial(seedBytes)
                                          random = New SecureRandom(digestGenerator)
                                      Else
                                          random = New SecureRandom()
                                      End If

                                      Dim keyGen As New RsaKeyPairGenerator()
                                      keyGen.Init(New KeyGenerationParameters(random, keySize))
                                      Dim keyPair As AsymmetricCipherKeyPair = keyGen.GenerateKeyPair()

                                      Dim privateKey As String
                                      Using sw As New StringWriter()
                                          Dim writer As New PemWriter(sw)
                                          writer.WriteObject(keyPair.Private)
                                          writer.Writer.Flush()
                                          privateKey = sw.ToString()
                                      End Using

                                      Dim publicKey As String
                                      Using sw As New StringWriter()
                                          Dim writer As New PemWriter(sw)
                                          writer.WriteObject(keyPair.Public)
                                          writer.Writer.Flush()
                                          publicKey = sw.ToString()
                                      End Using

                                      ' Global Variables for This Module
                                      RSAPublicKey = publicKey
                                      RSAPrivateKey = privateKey

                                      Return True
                                  Catch ex As Exception
                                      Debug.WriteLine("Error generating RSA keys: " & ex.Message)
                                      Return False
                                  End Try
                              End Function)
    End Function




    Public Sub RSA_ExportGeneratedKeyPairs()
        Dim keyPairsList As New List(Of RSAKeyPair)

        For Each outerKvp As KeyValuePair(Of Integer, Dictionary(Of String, String)) In RSAKeyPairs
            Dim keyPair As Dictionary(Of String, String) = outerKvp.Value

            ' Safely extract fields, using empty/default if not present
            Dim description As String = If(keyPair.ContainsKey("Description"), keyPair("Description"), "")
            Dim seed As String = If(keyPair.ContainsKey("Seed"), keyPair("Seed"), "")
            Dim salt As String = If(keyPair.ContainsKey("Salt"), keyPair("Salt"), "")
            Dim keySize As Integer = If(keyPair.ContainsKey("KeySize"), Integer.Parse(keyPair("KeySize")), 1024)

            ' Create a new RSAKeyPair object and add it to the list
            keyPairsList.Add(New RSAKeyPair With {
                .KeyIdentifier = outerKvp.Key,
                .PublicKey = keyPair("PublicKey"),
                .PrivateKey = keyPair("PrivateKey"),
                .Description = description,
                .Seed = seed,
                .Salt = salt,
                .KeySize = keySize
            })
        Next

        ' Serialize the list to a JSON string
        Dim json As String = JsonConvert.SerializeObject(keyPairsList, Formatting.Indented)

        ' Prompt the user to select a file to save the JSON
        Using saveFileDialog As New SaveFileDialog()
            saveFileDialog.Filter = "RSA Key Files|*.rsa|All files (*.*)|*.*"
            saveFileDialog.Title = "Save RSA Key Pairs"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Try
                    System.IO.File.WriteAllText(saveFileDialog.FileName, json)
                    MessageBox.Show("RSA Key Pairs exported successfully!", "Export Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error exporting key pairs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub




    Public Sub RSA_ImportGeneratedKeyPairs()
        Using openFileDialog As New OpenFileDialog()
            openFileDialog.Filter = "RSA Keys Files|*.rsa|All files (*.*)|*.*"
            openFileDialog.Title = "Open RSA Key Pairs"

            If openFileDialog.ShowDialog() = DialogResult.OK Then
                Try
                    Dim json As String = File.ReadAllText(openFileDialog.FileName)
                    Dim keyPairsList As List(Of RSAKeyPair) = JsonConvert.DeserializeObject(Of List(Of RSAKeyPair))(json)

                    For Each keyPair In keyPairsList
                        ' Create a new dictionary to hold the key pair data
                        Dim newKeyPair As New Dictionary(Of String, String) From {
                        {"PublicKey", keyPair.PublicKey},
                        {"PrivateKey", keyPair.PrivateKey},
                        {"KeySize", keyPair.KeySize.ToString()},
                        {"Seed", keyPair.Seed},
                        {"Salt", keyPair.Salt}
                    }

                        If Not String.IsNullOrEmpty(keyPair.Description) Then
                            newKeyPair.Add("Description", keyPair.Description)
                        End If

                        ' Add the key pair to the global RSAKeyPairs dictionary
                        RSAKeyPairs.Add(keyPair.KeyIdentifier, newKeyPair)
                    Next

                    MessageBox.Show("RSA Key Pairs imported successfully!", "Import Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error importing key pairs: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub






    Public Async Function RSA_ExtractLastCharsFromKey(publicOrPrivateKey As String, numOfChars As Integer) As Task(Of String)
        Dim tempString As String = Await RSA_ExtractKey_Public(publicOrPrivateKey)
        tempString = Await RSA_ExtractKey_Private(tempString)
        tempString = tempString.Trim()
        If tempString.Length > numOfChars Then
            Return tempString.Substring(tempString.Length - numOfChars)
        Else
            Return tempString
        End If
    End Function

    Public Async Function RSA_ExtractKey_Public(key As String) As Task(Of String)
        Return key.Replace("-----BEGIN PUBLIC KEY-----", "").Replace("-----END PUBLIC KEY-----", "").Replace(vbCrLf, "").Trim()
    End Function

    Public Async Function RSA_ExtractKey_Private(key As String) As Task(Of String)
        Return key.Replace("-----BEGIN RSA PRIVATE KEY-----", "").Replace("-----END RSA PRIVATE KEY-----", "").Replace(vbCrLf, "").Trim()
    End Function




    Public Async Function RsaEncrypt(plainText As String, publicKeyPem As String, salt As String) As Task(Of String)
        Return Await Task.Run(Function()
                                  Try
                                      Dim saltedPlainText As String = salt & plainText

                                      Dim publicKeyParam As AsymmetricKeyParameter
                                      Using reader As New StringReader(publicKeyPem)
                                          Dim pemReader As New PemReader(reader)
                                          publicKeyParam = CType(pemReader.ReadObject(), AsymmetricKeyParameter)
                                      End Using

                                      If publicKeyParam Is Nothing Then
                                          Console.WriteLine("empty key passed - generate a key pair first")
                                          MsgBox("empty key passed - generate a key pair first")
                                          Return String.Empty
                                      End If

                                      Dim encryptEngine As New RsaEngine()
                                      encryptEngine.Init(True, publicKeyParam)

                                      Dim inputBytes As Byte() = Encoding.UTF8.GetBytes(saltedPlainText)
                                      Dim encryptedBytes As Byte() = encryptEngine.ProcessBlock(inputBytes, 0, inputBytes.Length)

                                      Return Convert.ToBase64String(encryptedBytes)
                                  Catch ex As Exception
                                      Console.WriteLine("Encryption error: " & ex.Message)
                                      Return String.Empty
                                  End Try
                              End Function)
    End Function



    Public Async Function RsaDecrypt(cipherText As String, privateKeyPem As String, salt As String) As Task(Of String)
        Return Await Task.Run(Function()
                                  Try

                                      If String.IsNullOrEmpty(privateKeyPem) Then
                                          Console.WriteLine("empty key passed - generate a key pair first")
                                          MsgBox("empty key passed - generate a key pair first")
                                          Return String.Empty
                                      End If

                                      Dim privateKeyParam As AsymmetricKeyParameter
                                      Using reader As New StringReader(privateKeyPem)
                                          Dim pemReader As New PemReader(reader)
                                          privateKeyParam = CType(pemReader.ReadObject(), AsymmetricCipherKeyPair).Private
                                      End Using

                                      Dim decryptEngine As New RsaEngine()
                                      decryptEngine.Init(False, privateKeyParam)

                                      Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
                                      Dim decryptedBytes As Byte() = decryptEngine.ProcessBlock(cipherBytes, 0, cipherBytes.Length)

                                      Dim fullText As String = Encoding.UTF8.GetString(decryptedBytes)

                                      ' Remove salt if it's at the beginning
                                      If fullText.StartsWith(salt) Then
                                          Return fullText.Substring(salt.Length)
                                      Else
                                          Return fullText ' Fallback if no salt match
                                      End If
                                  Catch ex As Exception
                                      Console.WriteLine("Decryption error: " & ex.Message)
                                      Return String.Empty
                                  End Try
                              End Function)
    End Function



    Public Async Function RsaSignData(data As String, privateKeyPem As String, salt As String) As Task(Of String)
        Return Await Task.Run(Function()
                                  Try
                                      Dim saltedData = salt & data
                                      Dim keyPair As AsymmetricCipherKeyPair

                                      Using reader As New StringReader(privateKeyPem)
                                          Dim pemReader As New PemReader(reader)
                                          keyPair = CType(pemReader.ReadObject(), AsymmetricCipherKeyPair)
                                      End Using

                                      If privateKeyPem Is Nothing Then
                                          Console.WriteLine("empty key passed - generate a key pair first")
                                          MsgBox("empty key passed - generate a key pair first")
                                          Return String.Empty
                                      End If

                                      If keyPair Is Nothing Then
                                          Console.WriteLine("empty keyPair variable due to empty privateKeyPem - generate a key pair first")
                                          MsgBox("empty keyPair variable due to empty privateKeyPem - generate a key pair first")
                                          Return String.Empty
                                      End If

                                      Dim signer As ISigner = SignerUtilities.GetSigner("SHA256withRSA")
                                      signer.Init(True, keyPair.Private)

                                      Dim dataBytes = Encoding.UTF8.GetBytes(saltedData)
                                      signer.BlockUpdate(dataBytes, 0, dataBytes.Length)
                                      Dim signature As Byte() = signer.GenerateSignature()

                                      Return Convert.ToBase64String(signature)
                                  Catch ex As Exception
                                      Console.WriteLine("Signing error: " & ex.Message)
                                      Return String.Empty
                                  End Try
                              End Function)
    End Function



    Public Async Function RsaVerifySignature(data As String, signature As String, publicKeyPem As String, salt As String) As Task(Of Boolean)
        Return Await Task.Run(Function()
                                  Try
                                      Dim saltedData = salt & data

                                      Dim publicKeyParam As AsymmetricKeyParameter
                                      Using reader As New StringReader(publicKeyPem)
                                          Dim pemReader As New PemReader(reader)
                                          publicKeyParam = CType(pemReader.ReadObject(), AsymmetricKeyParameter)
                                      End Using

                                      If publicKeyParam Is Nothing Then
                                          Console.WriteLine("empty key passed - generate a key pair first")
                                          MsgBox("empty key passed - generate a key pair first")
                                          Return False
                                      End If

                                      Dim signer As ISigner = SignerUtilities.GetSigner("SHA256withRSA")
                                      signer.Init(False, publicKeyParam)

                                      Dim dataBytes As Byte() = Encoding.UTF8.GetBytes(saltedData)
                                      signer.BlockUpdate(dataBytes, 0, dataBytes.Length)

                                      Dim signatureBytes As Byte() = Convert.FromBase64String(signature)
                                      Return signer.VerifySignature(signatureBytes)
                                  Catch ex As Exception
                                      Console.WriteLine("Verification error: " & ex.Message)
                                      Return False
                                  End Try
                              End Function)
    End Function


    ' ----------------------------------












    ' #################
    ' HELPER FUNCTIONS
    ' #################



    Async Function GetSerials_HardDrivesAll() As Task(Of List(Of String))
        Try
            Dim query As New SelectQuery("SELECT SerialNumber FROM Win32_DiskDrive")
            Dim searcher As New ManagementObjectSearcher(query)
            Dim hardwareIds As New List(Of String)

            Await Task.Run(Sub()
                               For Each drive As ManagementObject In searcher.Get()
                                   Dim hardwareID As String = drive("SerialNumber").ToString().TrimEnd("."c)
                                   If Not String.IsNullOrEmpty(hardwareID) Then
                                       hardwareIds.Add(hardwareID)
                                   End If
                               Next
                           End Sub)

            Return hardwareIds

        Catch ex As Exception
            MessageBox.Show("An error occurred: " & ex.Message, "Error")
            Return New List(Of String)()
        End Try

    End Function






    Public Async Function GetAllDriveLettersAsync() As Task(Of List(Of String))
        Dim driveLetters As New List(Of String)()
        Await Task.Run(Async Function()
                           Dim allDrives As DriveInfo() = DriveInfo.GetDrives()
                           For Each drive As DriveInfo In allDrives
                               If drive.IsReady Then
                                   driveLetters.Add(drive.Name.Substring(0, 1))
                               End If
                           Next
                       End Function)

        Return driveLetters
    End Function









    ' V4
    Public Async Function GetDictionary_ComputerDetails() As Task(Of Dictionary(Of String, String))
        Return Await Task.Run(Function()
                                  Dim computerInfo As New Dictionary(Of String, String)

                                  ' Get the Hostname
                                  Try
                                      computerInfo("Hostname") = Dns.GetHostName()
                                  Catch ex As Exception
                                      computerInfo("Hostname") = "Error retrieving hostname: " & ex.Message
                                  End Try


                                  ' Get the Computer System information with specific columns
                                  Try
                                      ' Specify only the required properties in the SELECT statement
                                      Dim query As String = "SELECT Manufacturer, SystemFamily, Model, SystemSKUNumber, NumberOfProcessors, NumberOfLogicalProcessors, Domain, Workgroup, TotalPhysicalMemory FROM Win32_ComputerSystem"
                                      Dim managementClass As New ManagementObjectSearcher(query)

                                      For Each managementObject As ManagementObject In managementClass.Get()
                                          ' Extract the needed properties from the management object
                                          computerInfo("Make") = managementObject("Manufacturer").ToString()
                                          computerInfo("Model") = managementObject("SystemFamily").ToString()
                                          computerInfo("SystemModel") = managementObject("Model").ToString()
                                          computerInfo("SystemModelFull") = managementObject("SystemSKUNumber").ToString()
                                          computerInfo("CPUCount") = managementObject("NumberOfProcessors").ToString()
                                          computerInfo("CPUThreads") = managementObject("NumberOfLogicalProcessors").ToString()

                                          ' Check for Domain or Workgroup and add to dictionary
                                          If managementObject("Domain") IsNot Nothing Then
                                              computerInfo("Domain") = managementObject("Domain").ToString()
                                          Else
                                              computerInfo("Domain") = managementObject("Workgroup").ToString()
                                          End If

                                          ' Convert memory to GB and add to dictionary
                                          Dim totalMemory As ULong = Convert.ToUInt64(managementObject("TotalPhysicalMemory"))
                                          computerInfo("RAMSize") = Math.Round(totalMemory / (1024 * 1024 * 1024), 2).ToString() & " GB"
                                      Next
                                  Catch ex As Exception
                                      computerInfo("ComputerSystemError") = "Error retrieving system info: " & ex.Message
                                  End Try


                                  ' Get UUID 
                                  Try
                                      Dim uuidQuery As New ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct")
                                      For Each product As ManagementObject In uuidQuery.Get()
                                          computerInfo("UUID") = product("UUID").ToString()
                                      Next
                                  Catch ex As Exception
                                      computerInfo("Win32_ComputerSystemProductError") = "Error retrieving Win32_ComputerSystemProduct info: " & ex.Message
                                  End Try


                                  ' Get BIOS Information
                                  Try
                                      'Dim ramQuery As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory")
                                      Dim ramQuery As New ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS")
                                      For Each bios As ManagementObject In ramQuery.Get()
                                          computerInfo("BIOSSerial") = bios("SerialNumber").ToString()
                                      Next
                                  Catch ex As Exception
                                      computerInfo("BIOSSerialError") = "Error retrieving BIOS info: " & ex.Message
                                  End Try


                                  Try
                                      Dim motherBoardQuery As New ManagementObjectSearcher("SELECT SerialNumber, Product FROM Win32_BaseBoard")
                                      For Each queryResult As ManagementObject In motherBoardQuery.Get()
                                          Debug.WriteLine(queryResult.ToString)
                                          computerInfo("MotherboardSerial") = queryResult("SerialNumber").ToString()
                                          computerInfo("MotherBoardBaseBoardProduct") = queryResult("Product").ToString()
                                      Next
                                  Catch ex As Exception
                                      computerInfo("MotherboardError") = "Error retrieving Motherboard info: " & ex.Message
                                  End Try



                                  ' Get CPU Information
                                  Try
                                      'Dim cpuQuery As New ManagementObjectSearcher("SELECT * FROM Win32_Processor")
                                      Dim cpuQuery As New ManagementObjectSearcher("SELECT Name, MaxClockSpeed, NumberOfCores, Architecture FROM Win32_Processor")
                                      For Each cpu As ManagementObject In cpuQuery.Get()
                                          computerInfo("CPUName") = cpu("Name").ToString()

                                          ' Get CPU Speed in GHz
                                          Dim cpuSpeedMHz As Integer = Convert.ToInt32(cpu("MaxClockSpeed").ToString())
                                          Dim cpuSpeedGHz As Double = cpuSpeedMHz / 1000.0
                                          computerInfo("CPUSpeed") = cpuSpeedGHz.ToString("F2") & " GHz"

                                          computerInfo("CPUCores") = Convert.ToInt32(cpu("NumberOfCores").ToString())
                                          computerInfo("CPUArchitecture") = If(cpu("Architecture").ToString() = "9", "64-bit", "32-bit")
                                      Next
                                  Catch ex As Exception
                                      computerInfo("CPUError") = "Error retrieving CPU info: " & ex.Message
                                  End Try


                                  ' Get RAM Information
                                  Try
                                      'Dim ramQuery As New ManagementObjectSearcher("SELECT * FROM Win32_PhysicalMemory")
                                      Dim ramQuery As New ManagementObjectSearcher("SELECT Speed FROM Win32_PhysicalMemory")
                                      For Each ram As ManagementObject In ramQuery.Get()
                                          computerInfo("RAMSpeed") = ram("Speed").ToString() & " MHz"
                                      Next
                                  Catch ex As Exception
                                      computerInfo("RAMError") = "Error retrieving RAM info: " & ex.Message
                                  End Try


                                  ' Get OS Information
                                  Try
                                      'Dim osQuery As New ManagementObjectSearcher("SELECT * FROM Win32_OperatingSystem")
                                      Dim osQuery As New ManagementObjectSearcher("SELECT Caption, BuildNumber, Version, OSArchitecture FROM Win32_OperatingSystem")
                                      For Each os As ManagementObject In osQuery.Get()
                                          computerInfo("OSName") = os("Caption").ToString()
                                          computerInfo("OSBuild") = os("BuildNumber").ToString()
                                          computerInfo("OSVersion") = os("Version").ToString()
                                          computerInfo("OSArchitecture") = If(os("OSArchitecture").ToString() = "64-bit", "64-bit", "32-bit")

                                          ' Determine OS Service Channel (LTSB, LTSC, etc.)
                                          Dim caption As String = os("Caption").ToString()
                                          If caption.Contains("LTSB") Then
                                              computerInfo("OSServiceChannel") = "LTSB"
                                          ElseIf caption.Contains("LTSC") Then
                                              computerInfo("OSServiceChannel") = "LTSC"
                                          Else
                                              computerInfo("OSServiceChannel") = "Non-EnterpriseVersion"
                                          End If
                                      Next
                                  Catch ex As Exception
                                      computerInfo("OSError") = "Error retrieving OS info: " & ex.Message
                                  End Try

                                  ' Get Full Build Version from Registry
                                  Try
                                      Dim regKey As Microsoft.Win32.RegistryKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                                      If regKey IsNot Nothing Then
                                          Dim buildRevision As String = regKey.GetValue("UBR").ToString() ' UBR is the update build revision number
                                          computerInfo("OSBuild") = computerInfo("OSBuild").ToString() & "." & buildRevision
                                      End If
                                  Catch ex As Exception
                                      computerInfo("RegistryError") = "Error retrieving registry info: " & ex.Message
                                  End Try

                                  Return computerInfo

                              End Function)
    End Function






    Public Async Function GetDictionary_HardDriveDetails(Optional specificDriveLetter As String = "") As Task(Of Dictionary(Of String, String))
        ' By Default, extract the Drive letter from where the App is Running
        ' UNLESS a specific drive letter was provided as an argument to override it
        Dim driveLetter As String = ""
        If specificDriveLetter = "" Then
            ' Get the path where the application is running from
            Dim appPath As String = Application.StartupPath
            driveLetter = Path.GetPathRoot(appPath).TrimEnd("\"c)
        Else
            ' this way, doesnt matter is user passes "C", "C:", or "C:\", or even "C:\Anypath"
            driveLetter = specificDriveLetter.Substring(0, 1) & ":"
        End If

        Dim hardwareInfo As New Dictionary(Of String, String)
        Dim currentDriveHardwareID As String = ""

        'Dim diskQuery As New ManagementObjectSearcher(New ObjectQuery("SELECT * FROM Win32_DiskDrive"))
        Dim diskQuery As New ManagementObjectSearcher(New ObjectQuery("SELECT PNPDeviceID, DeviceID, Model, Size, InterfaceType, SerialNumber, Manufacturer FROM Win32_DiskDrive"))
        Dim diskNum As Integer = 1

        Try
            Dim disks As ManagementObjectCollection = Await Task.Run(Function() diskQuery.Get())

            hardwareInfo.Add("Count", disks.Count)

            For Each disk As ManagementObject In disks
                'Debug.WriteLine(vbCrLf & vbCrLf & "----------------------------------------" & vbCrLf & "Disk: " & diskNum & vbCrLf & "----------------------------------------")
                'Debug.WriteLine("Win32_DiskDrive --- Disk Information:")
                'Debug.WriteLine("PNPDeviceID: " & disk("PNPDeviceID"))
                'Debug.WriteLine("Disk DeviceID: " & disk("DeviceID"))
                'Debug.WriteLine("Model: " & disk("Model"))
                'Debug.WriteLine("Total Size: " & FormatSize(disk("Size")))
                'Debug.WriteLine("InterfaceType: " & disk("InterfaceType"))
                'Debug.WriteLine("SerialNumber: " & disk("SerialNumber"))
                '''''Debug.WriteLine("Manufacturer: " & disk("Manufacturer"))
                '''''Debug.WriteLine("MediaType: " & disk("MediaType"))
                '''''Debug.WriteLine("Status: " & disk("Status"))
                '''''Debug.WriteLine("FirmwareRevision: " & disk("FirmwareRevision"))
                '''''Debug.WriteLine()

                Dim partitionQuery As New ManagementObjectSearcher(New ObjectQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID='" & disk("DeviceID") & "'} WHERE AssocClass=Win32_DiskDriveToDiskPartition"))
                Dim partitions As ManagementObjectCollection = Await Task.Run(Function() partitionQuery.Get())

                'Debug.WriteLine("PARTITION TOTAL COUNT: " & partitions.Count)

                For Each partition As ManagementObject In partitions

                    Dim logicalDiskQuery As New ManagementObjectSearcher(New ObjectQuery("ASSOCIATORS OF {Win32_DiskPartition.DeviceID='" & partition("DeviceID") & "'} WHERE AssocClass=Win32_LogicalDiskToPartition"))
                    Dim logicalDisks As ManagementObjectCollection = Await Task.Run(Function() logicalDiskQuery.Get())

                    For Each logicalDisk As ManagementObject In logicalDisks
                        'Debug.WriteLine("")
                        'Debug.WriteLine("DriveLetter: " & logicalDisk("DeviceID"))
                        'Debug.WriteLine("VolumeName: " & If(logicalDisk("VolumeName") IsNot Nothing, logicalDisk("VolumeName").ToString(), "N/A"))
                        'Debug.WriteLine("VolumeSerialNumber: " & logicalDisk("VolumeSerialNumber"))
                        'Debug.WriteLine("Size: " & FormatSize(logicalDisk("Size")))
                        'Debug.WriteLine("FreeSpace: " & FormatSize(logicalDisk("FreeSpace")))
                        '''''Debug.WriteLine("Win32_DiskPartition --- Logical Disk Information:")
                        '''''Debug.WriteLine("DriveType: " & logicalDisk("DriveType"))
                        '''''Debug.WriteLine("FileSystem: " & logicalDisk("FileSystem"))
                        '''''Debug.WriteLine("Compressed: " & logicalDisk("Compressed"))

                        If logicalDisk("DeviceID") = driveLetter Then
                            hardwareInfo.Add("PNPDeviceID", disk("PNPDeviceID"))
                            hardwareInfo.Add("VolumeSerialNumber", logicalDisk("VolumeSerialNumber"))
                            hardwareInfo.Add("DriveSerialNumber", disk("SerialNumber").ToString().TrimEnd("."c))
                            hardwareInfo.Add("Model", disk("Model"))
                            hardwareInfo.Add("DiskDeviceID", disk("DeviceID"))
                            hardwareInfo.Add("SizeB", disk("Size").ToString())
                            hardwareInfo.Add("SizeGB", FormatSize(disk("Size")))
                            hardwareInfo.Add("FreeSpaceB", logicalDisk("FreeSpace"))
                            hardwareInfo.Add("FreeSpaceGB", FormatSize(logicalDisk("FreeSpace")))
                            hardwareInfo.Add("DriveLetter", logicalDisk("DeviceID"))
                            hardwareInfo.Add("VolumeName", If(logicalDisk("VolumeName") IsNot Nothing, logicalDisk("VolumeName").ToString(), "N/A"))
                            hardwareInfo.Add("FileSystem", logicalDisk("FileSystem"))
                        End If
                    Next
                Next
                diskNum += 1
            Next
        Catch ex As Exception
            MessageBox.Show("Error obtaining drive information" & vbCrLf & vbCrLf & ex.Message & vbCrLf & vbCrLf & ex.StackTrace, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try

        Return hardwareInfo
    End Function







    Public Async Function GetDictionary_NetworkDetails() As Task(Of Dictionary(Of String, Dictionary(Of String, String)))
        ' The result dictionary where each key is an adapter's name or identifier, and the value is another dictionary with adapter information
        Dim networkInfo As New Dictionary(Of String, Dictionary(Of String, String))()

        Try
            ' Run the task asynchronously to avoid blocking
            Await Task.Run(Sub()
                               ' Get the network interfaces on the local machine
                               Dim networkInterfaces As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces()

                               ' Loop through each network interface (adapter)
                               For Each netInterface As NetworkInterface In networkInterfaces
                                   Try
                                       Dim adapterInfo As New Dictionary(Of String, String)()

                                       adapterInfo("NICDescription") = netInterface.Description
                                       adapterInfo("MACAddress") = netInterface.GetPhysicalAddress().ToString()

                                       Dim ipProperties As IPInterfaceProperties = netInterface.GetIPProperties()
                                       Dim ipv4Address As String = String.Empty
                                       Dim subnetMask As String = String.Empty
                                       Dim defaultGateway As String = String.Empty

                                       ' Find the first IPv4 address and its subnet mask
                                       For Each unicastAddr As UnicastIPAddressInformation In ipProperties.UnicastAddresses
                                           If unicastAddr.Address.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                                               ipv4Address = unicastAddr.Address.ToString()
                                               subnetMask = unicastAddr.IPv4Mask.ToString()
                                               Exit For
                                           End If
                                       Next

                                       ' Add IPv4 Address and Subnet Mask to the dictionary only if found
                                       If Not String.IsNullOrEmpty(ipv4Address) Then
                                           adapterInfo("IPv4Address") = ipv4Address
                                           adapterInfo("SubnetMask") = subnetMask
                                       End If

                                       ' Get the default gateway (ensure it's IPv4)
                                       For Each gateway As GatewayIPAddressInformation In ipProperties.GatewayAddresses
                                           If gateway.Address.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                                               defaultGateway = gateway.Address.ToString()
                                               Exit For
                                           End If
                                       Next

                                       ' Add Default Gateway to the dictionary only if found
                                       If Not String.IsNullOrEmpty(defaultGateway) Then
                                           adapterInfo("DefaultGateway") = defaultGateway
                                       End If

                                       ' Get the DNS domain suffix (if available)
                                       Dim ipPropertiesGlobal As IPGlobalProperties = IPGlobalProperties.GetIPGlobalProperties()
                                       Dim dnsDomainSuffix As String = String.Empty

                                       If ipPropertiesGlobal.DomainName IsNot Nothing Then
                                           'Debug.WriteLine("ipPropertiesGlobal.DomainName = " & ipPropertiesGlobal.DomainName)
                                           dnsDomainSuffix = ipPropertiesGlobal.DomainName
                                       End If

                                       'Debug.WriteLine("ipProperties.DnsSuffix = " & ipProperties.DnsSuffix)
                                       If String.IsNullOrEmpty(dnsDomainSuffix) = True Then
                                           If String.IsNullOrEmpty(ipProperties.DnsSuffix) = False Then
                                               dnsDomainSuffix = ipProperties.DnsSuffix
                                           End If
                                       End If

                                       adapterInfo("DNSDomainSuffix") = dnsDomainSuffix

                                       ' Add the adapter information to the result dictionary
                                       networkInfo(netInterface.Name) = adapterInfo

                                   Catch ex As Exception
                                       ' Handle any errors specific to a single network interface
                                       Debug.WriteLine("Error processing adapter " & netInterface.Name & ": " & ex.Message)
                                   End Try
                               Next
                           End Sub)
        Catch ex As Exception
            Debug.WriteLine("Error retrieving network details: " & ex.Message)
        End Try

        Return networkInfo
    End Function






    ' Enhanced helper function to validate an adapter 
    Private Function IsValidAdapter(ByVal adapterName As String, ByVal NICDescription As String) As Boolean
        ' Convert both to lowercase for case-insensitive comparison
        Dim tempAdapterName As String = adapterName.ToLower
        Dim tempNICDescription As String = NICDescription.ToLower

        ' exclude virtual, loopback, Bluetooth, and virtualization software
        If tempAdapterName.Contains("loopback") OrElse tempNICDescription.Contains("loopback") OrElse tempAdapterName.Contains("virtual") OrElse tempNICDescription.Contains("virtual") Then
            Return False
        End If

        If tempAdapterName.Contains("bluetooth") OrElse tempNICDescription.Contains("bluetooth") Then
            Return False
        End If

        If tempAdapterName.Contains("vmware") OrElse tempNICDescription.Contains("vmware") OrElse tempAdapterName.Contains("virtualbox") OrElse tempNICDescription.Contains("virtualbox") OrElse tempAdapterName.Contains("hyper-v") OrElse tempNICDescription.Contains("hyper-v") OrElse tempAdapterName.Contains("parallels") OrElse tempNICDescription.Contains("parallels") OrElse tempAdapterName.Contains("docker") OrElse tempNICDescription.Contains("docker") Then
            Return False
        End If

        Return True
    End Function





    Private Function FormatSize(ByVal size As Object) As String
        If size IsNot Nothing Then
            Dim bytes As Long = Convert.ToInt64(size)
            If bytes >= 1L * 1024 * 1024 * 1024 Then
                Return (bytes / (1024 * 1024 * 1024)).ToString("F2") & " GB"
            ElseIf bytes >= 1L * 1024 * 1024 Then
                Return (bytes / (1024 * 1024)).ToString("F2") & " MB"
            ElseIf bytes >= 1L * 1024 Then
                Return (bytes / 1024).ToString("F2") & " KB"
            Else
                Return bytes & " Bytes"
            End If
        Else
            Return "N/A"
        End If
    End Function




    Public Sub CopyStringToClipboard(text As String)
        Try
            Clipboard.SetText(text)
        Catch ex As Exception
            MessageBox.Show("An error occurred while copying to clipboard: " & ex.Message)
        End Try
    End Sub





    Public Function Base64String_Encode(passedString As String) As String
        Dim inputString As String = passedString.Trim()
        If Not String.IsNullOrEmpty(inputString) Then
            Dim inputBytes As Byte() = Encoding.UTF8.GetBytes(inputString)
            Dim base64Encoded As String = Convert.ToBase64String(inputBytes)
            Return base64Encoded
        Else
            MessageBox.Show("Please enter a string to encode.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return String.Empty
        End If
    End Function




    Public Function Base64String_Decode(passedString As String) As String
        Dim base64EncodedString As String = passedString.Trim()
        If Not String.IsNullOrEmpty(base64EncodedString) Then
            Try
                Dim decodedBytes As Byte() = Convert.FromBase64String(base64EncodedString)
                Dim decodedString As String = Encoding.UTF8.GetString(decodedBytes)
                Return decodedString
            Catch ex As FormatException
                MessageBox.Show("Invalid Base64 string format. Please check the input.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return String.Empty
            End Try
        Else
            MessageBox.Show("Please enter a Base64 string to decode.", "Input Required", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return String.Empty
        End If
    End Function




    Function SplitAndReverseString(passedString As String) As String
        ' Identify padding at the end of the string
        Dim padding As String = ""
        Dim trimmedBase64 As String = passedString

        If passedString.EndsWith("=") Then
            If passedString.EndsWith("==") Then
                ' If the string ends with '==', trim the last 2 characters
                padding = "=="
                trimmedBase64 = passedString.Substring(0, passedString.Length - 2)
            ElseIf passedString.EndsWith("=") Then
                ' If the string ends with '=', trim the last 1 character
                padding = "="
                trimmedBase64 = passedString.Substring(0, passedString.Length - 1)
            End If
        End If

        ' Calculate the middle point for splitting the string
        Dim midPoint As Integer = trimmedBase64.Length \ 2

        ' Split the string into two parts
        Dim string1 As String = trimmedBase64.Substring(0, midPoint)
        Dim string2 As String = trimmedBase64.Substring(midPoint)

        ' Reverse the order and append the padding if present
        Return string2 & string1 & padding
    End Function




    ' Simple Base32 encoder/decoder
    Private Const Base32Chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"



    Public Function Base32String_Encode(byteData As Byte()) As String
        Dim output As New StringBuilder()
        Dim bitBuffer As Integer = 0
        Dim bitCount As Integer = 0

        For Each b In byteData
            bitBuffer = (bitBuffer << 8) Or b
            bitCount += 8

            While bitCount >= 5
                Dim index = (bitBuffer >> (bitCount - 5)) And 31
                output.Append(Base32Chars(index))
                bitCount -= 5
            End While
        Next

        If bitCount > 0 Then
            Dim index = (bitBuffer << (5 - bitCount)) And 31
            output.Append(Base32Chars(index))
        End If

        Return output.ToString()
    End Function



    Public Function Base32String_Decode(base32 As String) As Byte()
        Dim bitBuffer As Integer = 0
        Dim bitCount As Integer = 0
        Dim output As New List(Of Byte)()

        For Each c In base32.ToUpper()
            If Not Base32Chars.Contains(c) Then Continue For
            bitBuffer = (bitBuffer << 5) Or Base32Chars.IndexOf(c)
            bitCount += 5

            If bitCount >= 8 Then
                output.Add(CByte((bitBuffer >> (bitCount - 8)) And 255))
                bitCount -= 8
            End If
        Next

        Return output.ToArray()
    End Function



    Public Function GenerateUUID_LicGen() As String
        Return Guid.NewGuid().ToString()
    End Function


    ' Time-Based OTP
    Public Function GenerateTOTP(base32Secret As String, Optional stepSeconds As Integer = 30) As String
        Try
            Dim secretBytes As Byte() = Base32String_Decode(base32Secret)
            Dim totp As New Totp(secretBytes, stepSeconds, OtpHashMode.Sha1, 6)
            Return totp.ComputeTotp()
        Catch ex As Exception
            MessageBox.Show("TOTP generation error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function


    ' Counter-Based OTP
    Public Function GenerateHOTP(base32Secret As String, counter As Long) As String
        Try
            Dim secretBytes As Byte() = Base32String_Decode(base32Secret)
            Dim hotp As New Hotp(secretBytes, OtpHashMode.Sha1, 6)
            Return hotp.ComputeHOTP(counter)
        Catch ex As Exception
            MessageBox.Show("HOTP generation error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return String.Empty
        End Try
    End Function


End Module

```


