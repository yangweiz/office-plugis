using System;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;

namespace Suite.Infrastructure
{
    public sealed class JsonSettingsStore<T> where T : class, new()
    {
        private readonly string _path;
        public JsonSettingsStore(string path) { if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("设置路径不能为空。", "path"); _path = path; }
        public string LastError { get; private set; }

        public T Load()
        {
            LastError = null;
            if (!File.Exists(_path)) return new T();
            try
            {
                using (var stream = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    var serializer = new DataContractJsonSerializer(typeof(T));
                    return serializer.ReadObject(stream) as T ?? new T();
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Runtime.Serialization.SerializationException)
            {
                LastError = "设置读取失败，已使用安全默认值；原设置文件未修改。" + ex.Message;
                return new T();
            }
        }

        public void Save(T settings)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            string directory = Path.GetDirectoryName(_path);
            Directory.CreateDirectory(directory);
            string temporary = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    var serializer = new DataContractJsonSerializer(typeof(T));
                    serializer.WriteObject(stream, settings);
                    stream.Flush(true);
                }
                if (File.Exists(_path)) File.Replace(temporary, _path, _path + ".bak");
                else File.Move(temporary, _path);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
            }
        }
    }
}
