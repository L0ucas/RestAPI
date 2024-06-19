using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security.Cryptography;
using System.IO;

namespace TMHelper.Common
{
    public static class CryptographyHelper
    {
        #region Constant
        private const string _passwordHash = "@cc1234$$H@$h";
        private const string _saltKey = "@cc1234$$$@ltK3y";
        private const string _vIKey = "@cc1234$$V1K3y!G";
        #endregion

        public static string Encrypt(string plainText)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            byte[] keyBytes = new Rfc2898DeriveBytes(_passwordHash, Encoding.ASCII.GetBytes(_saltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
            var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(_vIKey));

            byte[] cipherTextBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                    cryptoStream.FlushFinalBlock();
                    cipherTextBytes = memoryStream.ToArray();
                    cryptoStream.Close();
                }
                memoryStream.Close();
            }
            return Convert.ToBase64String(cipherTextBytes);
        }

        public static string Encrypt(string plainText, string strPasswordHash, string strSaltKey)
        {
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(plainText);

            byte[] keyBytes = new Rfc2898DeriveBytes(strPasswordHash, Encoding.ASCII.GetBytes(strSaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.Zeros };
            var encryptor = symmetricKey.CreateEncryptor(keyBytes, Encoding.ASCII.GetBytes(_vIKey));

            byte[] cipherTextBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write))
                {
                    cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
                    cryptoStream.FlushFinalBlock();
                    cipherTextBytes = memoryStream.ToArray();
                    cryptoStream.Close();
                }
                memoryStream.Close();
            }
            return Convert.ToBase64String(cipherTextBytes);
        }

        public static string Decrypt(string encryptedText)
        {
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
            byte[] keyBytes = new Rfc2898DeriveBytes(_passwordHash, Encoding.ASCII.GetBytes(_saltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

            var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(_vIKey));
            var memoryStream = new MemoryStream(cipherTextBytes);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
        }

        public static string Decrypt(string encryptedText, string strPasswordHash, string strSaltKey)
        {
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedText);
            byte[] keyBytes = new Rfc2898DeriveBytes(strPasswordHash, Encoding.ASCII.GetBytes(strSaltKey)).GetBytes(256 / 8);
            var symmetricKey = new RijndaelManaged() { Mode = CipherMode.CBC, Padding = PaddingMode.None };

            var decryptor = symmetricKey.CreateDecryptor(keyBytes, Encoding.ASCII.GetBytes(_vIKey));
            var memoryStream = new MemoryStream(cipherTextBytes);
            var cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];

            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount).TrimEnd("\0".ToCharArray());
        }

        public static string MD5Encrypt(string pToEncrypt, string sKey = "??Cv?r?8")
        {
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            byte[] inputByteArray = Encoding.Default.GetBytes(pToEncrypt);
            des.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
            des.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            StringBuilder ret = new StringBuilder();
            foreach (byte b in ms.ToArray())
            {
                ret.AppendFormat("{0:X2}", b);
            }
            ret.ToString();
            return ret.ToString();
        }

        public static string MD5Decrypt(string pToDecrypt, string sKey = "??Cv?r?8")
        {
            if (string.IsNullOrEmpty(pToDecrypt))
            {
                return "";
            }

            DESCryptoServiceProvider des = new DESCryptoServiceProvider();

            byte[] inputByteArray = new byte[pToDecrypt.Length / 2];
            for (int x = 0; x < pToDecrypt.Length / 2; x++)
            {
                int i = (Convert.ToInt32(pToDecrypt.Substring(x * 2, 2), 16));
                inputByteArray[x] = (byte)i;
            }

            des.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
            des.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();

            StringBuilder ret = new StringBuilder();

            return System.Text.Encoding.Default.GetString(ms.ToArray());
        }

        public static string SHA512Encrypt(string strToEncrypt, string strHashKey = _passwordHash, string strSaltKey = _saltKey)
        {
            string strOutputEncrypt = string.Empty;

            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strToEncrypt))
                    return string.Empty;
                if (string.IsNullOrWhiteSpace(strHashKey))
                    return string.Empty;
                if (string.IsNullOrWhiteSpace(strSaltKey))
                    return string.Empty;
                #endregion

                byte[] bToEncrypt = Encoding.UTF8.GetBytes(strToEncrypt);
                byte[] bHashKey = Encoding.UTF8.GetBytes(strHashKey);
                byte[] bSaltKey = Encoding.UTF8.GetBytes(strSaltKey);
                byte[] bSHA512ComputeHash = SHA512.Create().ComputeHash(bHashKey);

                byte[] bOutputEncrypted = Encrypt(bToEncrypt, bSHA512ComputeHash, bSaltKey);
                strOutputEncrypt = Convert.ToBase64String(bOutputEncrypted);
            }
            catch (Exception ex)
            {
                throw ex; 
            }

            return strOutputEncrypt;
        }

        public static string SHA512Decrypt(string strEncypted, string strHashKey = _passwordHash, string strSaltKey = _saltKey)
        {
            string strOutputDecrypt = string.Empty;
            try
            {
                #region Pre-checking
                if (string.IsNullOrWhiteSpace(strEncypted))
                    return string.Empty;
                if (string.IsNullOrWhiteSpace(strHashKey))
                    return string.Empty;
                if (string.IsNullOrWhiteSpace(strSaltKey))
                    return string.Empty;
                #endregion

                byte[] bEncrypted = Convert.FromBase64String(strEncypted);
                byte[] bHashKey = Encoding.UTF8.GetBytes(strHashKey);
                byte[] bSaltKey = Encoding.UTF8.GetBytes(strSaltKey);
                byte[] bEncryptedVal = SHA512.Create().ComputeHash(bHashKey);

                var bOutputDecrypt = Decrypt(bEncrypted, bEncryptedVal, bSaltKey);
                strOutputDecrypt = Encoding.UTF8.GetString(bOutputDecrypt);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return strOutputDecrypt;
        }

        public static string GenRandomAlphaNumString()
        {
            string strAlphaNum = string.Empty;

            try
            {
                List<string> lstLAlphabet = new List<string> { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
                List<string> lstUAlphabet = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                List<string> lstNumberic = new List<string> { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", "23", "24", "25", "26" };

                Random random = new Random();
                int iLAIndex = random.Next(0, 25);
                int iUAIndex = random.Next(0, 25);
                int iNIndex = random.Next(0, 25);

                string strGuid = Guid.NewGuid().ToString().Replace("-", string.Empty);
                strAlphaNum = string.Format("{0}@{1}{2}{3}", strGuid.Substring(strGuid.Length - 6), lstLAlphabet[iLAIndex], lstUAlphabet[iUAIndex], lstNumberic[iNIndex]);
            }
            catch (Exception ex) {}

            return strAlphaNum;
        }

        #region Helper
        private static byte[] Encrypt(byte[] bytesToBeEncrypted, byte[] passwordBytes, byte[] saltBytes)
        {
            byte[] encryptedBytes = null;

            using (MemoryStream ms = new MemoryStream())
            {
                using (RijndaelManaged AES = new RijndaelManaged())
                {
                    var key = new Rfc2898DeriveBytes(passwordBytes, saltBytes, 1000);

                    AES.KeySize = 256;
                    AES.BlockSize = 128;
                    AES.Key = key.GetBytes(AES.KeySize / 8);
                    AES.IV = key.GetBytes(AES.BlockSize / 8);

                    AES.Mode = CipherMode.CBC;

                    using (var cs = new CryptoStream(ms, AES.CreateEncryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytesToBeEncrypted, 0, bytesToBeEncrypted.Length);
                        cs.Close();
                    }

                    encryptedBytes = ms.ToArray();
                }
            }

            return encryptedBytes;
        }

        private static byte[] Decrypt(byte[] bytesToBeDecrypted, byte[] passwordBytes, byte[] saltBytes)
        {
            byte[] decryptedBytes = null;

            using (MemoryStream ms = new MemoryStream())
            {
                using (RijndaelManaged AES = new RijndaelManaged())
                {
                    var key = new Rfc2898DeriveBytes(passwordBytes, saltBytes, 1000);

                    AES.KeySize = 256;
                    AES.BlockSize = 128;
                    AES.Key = key.GetBytes(AES.KeySize / 8);
                    AES.IV = key.GetBytes(AES.BlockSize / 8);
                    AES.Mode = CipherMode.CBC;

                    using (var cs = new CryptoStream(ms, AES.CreateDecryptor(), CryptoStreamMode.Write))
                    {
                        cs.Write(bytesToBeDecrypted, 0, bytesToBeDecrypted.Length);
                        cs.Close();
                    }

                    decryptedBytes = ms.ToArray();
                }
            }

            return decryptedBytes;
        }
        #endregion
    }
}
