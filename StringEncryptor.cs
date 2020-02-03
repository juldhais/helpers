using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Helpers
{
    public static class StringEncryptor
    {
        public static string Encrypt(string text, string password = "123456789")
        {
            try
            {
                byte[] Results;
                UTF8Encoding UTF8 = new UTF8Encoding();

                using (MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider())
                {
                    byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(password));

                    using (TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider())
                    {
                        TDESAlgorithm.Key = TDESKey;
                        TDESAlgorithm.Mode = CipherMode.ECB;
                        TDESAlgorithm.Padding = PaddingMode.PKCS7;
                        byte[] DataToEncrypt = UTF8.GetBytes(text);

                        try
                        {
                            ICryptoTransform Encryptor = TDESAlgorithm.CreateEncryptor();
                            Results = Encryptor.TransformFinalBlock(DataToEncrypt, 0, DataToEncrypt.Length);
                        }
                        finally
                        {
                            TDESAlgorithm.Clear();
                            HashProvider.Clear();
                        }
                    }
                }

                return Convert.ToBase64String(Results);
            }
            catch
            {
                return text;
            }
        }

        public static string Decrypt(string text, string password = "123456789")
        {
            try
            {
                byte[] Results;
                UTF8Encoding UTF8 = new UTF8Encoding();

                using (MD5CryptoServiceProvider HashProvider = new MD5CryptoServiceProvider())
                {
                    byte[] TDESKey = HashProvider.ComputeHash(UTF8.GetBytes(password));

                    using (TripleDESCryptoServiceProvider TDESAlgorithm = new TripleDESCryptoServiceProvider())
                    {
                        TDESAlgorithm.Key = TDESKey;
                        TDESAlgorithm.Mode = CipherMode.ECB;
                        TDESAlgorithm.Padding = PaddingMode.PKCS7;
                        byte[] DataToDecrypt = Convert.FromBase64String(text);
                        try
                        {
                            ICryptoTransform Decryptor = TDESAlgorithm.CreateDecryptor();
                            Results = Decryptor.TransformFinalBlock(DataToDecrypt, 0, DataToDecrypt.Length);
                        }
                        finally
                        {
                            TDESAlgorithm.Clear();
                            HashProvider.Clear();
                        }
                    }
                }

                return UTF8.GetString(Results);
            }
            catch
            {
                return text;
            }
        }
    }
}
