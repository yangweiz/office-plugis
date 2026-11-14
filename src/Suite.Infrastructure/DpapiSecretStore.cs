using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Suite.Infrastructure
{
    public sealed class DpapiSecretStore
    {
        private readonly string _directory;
        public DpapiSecretStore(string directory) { _directory = directory ?? throw new ArgumentNullException("directory"); }

        public void Save(string name, string secret)
        {
            ValidateName(name);
            if (secret == null) throw new ArgumentNullException("secret");
            Directory.CreateDirectory(_directory);
            byte[] encrypted = ProtectedData.Protect(Encoding.UTF8.GetBytes(secret), null, DataProtectionScope.CurrentUser);
            string path = GetPath(name);
            string temporary = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllText(temporary, Convert.ToBase64String(encrypted), Encoding.ASCII);
            if (File.Exists(path)) File.Replace(temporary, path, null); else File.Move(temporary, path);
        }

        public bool TryRead(string name, out string secret)
        {
            ValidateName(name);
            secret = null;
            string path = GetPath(name);
            if (!File.Exists(path)) return false;
            try
            {
                byte[] decrypted = ProtectedData.Unprotect(Convert.FromBase64String(File.ReadAllText(path, Encoding.ASCII)), null, DataProtectionScope.CurrentUser);
                secret = Encoding.UTF8.GetString(decrypted);
                Array.Clear(decrypted, 0, decrypted.Length);
                return true;
            }
            catch (Exception ex) when (ex is CryptographicException || ex is FormatException || ex is IOException || ex is UnauthorizedAccessException) { return false; }
        }

        public void Delete(string name)
        {
            ValidateName(name);
            string path = GetPath(name);
            if (File.Exists(path)) File.Delete(path);
        }

        private string GetPath(string name)
        {
            using (var sha = SHA256.Create()) return Path.Combine(_directory, Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(name))).Replace('/', '_').Replace('+', '-').TrimEnd('=') + ".secret");
        }

        private static void ValidateName(string name) { if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("凭据名称不能为空。", "name"); }
    }
}
