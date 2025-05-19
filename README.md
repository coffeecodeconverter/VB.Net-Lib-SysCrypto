# VB.Net-Lib-SysCrypto
VB.Net Module for System, Network, and User Info, plus Encryption, and Encoding functions - can be used for a software licensing system

# 🔐 VB.NET System Info, Encryption & Licensing Module

A powerful and versatile VB.NET module for gathering system, user, network, and hardware details, while offering strong cryptographic tools such as AES/RSA encryption, hashing, UUID generation, and OTP (TOTP/HOTP) support.

Originally created for a flexible **software licensing system**, but useful across many domains including telemetry, authentication, diagnostics, and hardware binding.

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



🔑 RSA Cryptography

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


🧠 System Information

Await GetThe_Computer("Hostname")
Await GetThe_Computer("CPUName")
Await GetThe_HardDrive("DriveSerialNumber")
Await GetThe_Network("IPv4Address")
Await GetThe_User("Username")
Await GetFingerprint_Computer()


