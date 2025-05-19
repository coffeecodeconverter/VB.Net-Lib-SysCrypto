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
