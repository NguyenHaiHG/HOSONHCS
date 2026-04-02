using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace HOSONHCS
{
    /// <summary>
    /// Mã hóa / giải mã toanquoc.json bằng AES-256.
    /// File plaintext (.json) chỉ tồn tại lúc đóng gói, không gửi cho user.
    /// User chỉ nhận toanquoc.enc (bytes rác, không đọc được).
    /// </summary>
    internal static class ToanQuocEncryptor
    {
        private const string ENC_FILE = "toanquoc.enc";
        private const string JSON_FILE = "toanquoc.json";

        // Cache nội dung đã giải mã trong RAM — không ghi ra file
        private static string _decryptedJson = null;
        private static readonly object _lock = new object();

        /// <summary>
        /// Lấy nội dung JSON (giải mã từ .enc hoặc đọc .json rồi tự mã hóa).
        /// Kết quả được cache trong RAM.
        /// </summary>
        public static string GetJson()
        {
            lock (_lock)
            {
                if (_decryptedJson != null) return _decryptedJson;

                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var encPath  = Path.Combine(baseDir, ENC_FILE);
                var jsonPath = Path.Combine(baseDir, JSON_FILE);

                // Ưu tiên đọc file .enc
                if (File.Exists(encPath))
                {
                    _decryptedJson = Decrypt(File.ReadAllBytes(encPath));
                    return _decryptedJson;
                }

                // Lần đầu: có file .json plaintext → mã hóa → xóa .json
                if (File.Exists(jsonPath))
                {
                    var json = File.ReadAllText(jsonPath, Encoding.UTF8);
                    var encrypted = Encrypt(json);
                    File.WriteAllBytes(encPath, encrypted);
                    File.Delete(jsonPath);          // xóa plaintext
                    _decryptedJson = json;
                    return _decryptedJson;
                }

                return null;
            }
        }

        /// <summary>
        /// Ghi lại nội dung JSON (sau khi chỉnh sửa qua app) → mã hóa lại .enc.
        /// </summary>
        public static void SaveJson(string json)
        {
            lock (_lock)
            {
                var encPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ENC_FILE);
                File.WriteAllBytes(encPath, Encrypt(json));
                _decryptedJson = json;
            }
        }

        /// <summary>Xóa cache RAM (gọi khi cần reload).</summary>
        public static void ClearCache() { lock (_lock) { _decryptedJson = null; } }

        /// <summary>
        /// Xuất nội dung hiện tại ra file .json để chỉnh sửa bên ngoài.
        /// Sau khi chỉnh sửa xong, gọi ImportFromJson để mã hóa lại.
        /// </summary>
        public static bool ExportToJson(string outputJsonPath)
        {
            var json = GetJson();
            if (json == null) return false;
            File.WriteAllText(outputJsonPath, json, Encoding.UTF8);
            return true;
        }

        /// <summary>
        /// Đọc file .json đã chỉnh sửa, mã hóa lại thành .enc và xóa file .json.
        /// </summary>
        public static bool ImportFromJson(string inputJsonPath)
        {
            if (!File.Exists(inputJsonPath)) return false;
            var json = File.ReadAllText(inputJsonPath, Encoding.UTF8);
            SaveJson(json);
            File.Delete(inputJsonPath);
            TinhHelper.RefreshCache();
            return true;
        }

        // ── AES Encrypt ──────────────────────────────────────────────────────
        private static byte[] Encrypt(string plainText)
        {
            using (var aes = Aes.Create())
            {
                aes.Key  = AppSecrets.AesKey;
                aes.IV   = AppSecrets.AesIV;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                using (var enc = aes.CreateEncryptor())
                using (var ms  = new MemoryStream())
                using (var cs  = new CryptoStream(ms, enc, CryptoStreamMode.Write))
                {
                    var bytes = Encoding.UTF8.GetBytes(plainText);
                    cs.Write(bytes, 0, bytes.Length);
                    cs.FlushFinalBlock();
                    return ms.ToArray();
                }
            }
        }

        // ── AES Decrypt ──────────────────────────────────────────────────────
        private static string Decrypt(byte[] cipherBytes)
        {
            using (var aes = Aes.Create())
            {
                aes.Key  = AppSecrets.AesKey;
                aes.IV   = AppSecrets.AesIV;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                using (var dec = aes.CreateDecryptor())
                using (var ms  = new MemoryStream(cipherBytes))
                using (var cs  = new CryptoStream(ms, dec, CryptoStreamMode.Read))
                using (var sr  = new StreamReader(cs, Encoding.UTF8))
                {
                    return sr.ReadToEnd();
                }
            }
        }
    }
}
