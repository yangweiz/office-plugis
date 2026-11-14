using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Suite.Infrastructure
{
    public enum LocalContentKind { MacroSnippet = 1, Inspiration = 2 }

    [DataContract]
    public sealed class LocalContentItem
    {
        [DataMember(Order = 1)] public string Id { get; set; }
        [DataMember(Order = 2)] public LocalContentKind Kind { get; set; }
        [DataMember(Order = 3)] public string Title { get; set; }
        [DataMember(Order = 4)] public string Content { get; set; }
        [DataMember(Order = 5)] public string Tags { get; set; }
        [DataMember(Order = 6)] public DateTime CreatedUtc { get; set; }
        [DataMember(Order = 7)] public DateTime UpdatedUtc { get; set; }
    }

    [DataContract]
    public sealed class LocalContentData
    {
        public LocalContentData() { SchemaVersion = 1; Items = new List<LocalContentItem>(); }
        [DataMember(Order = 1)] public int SchemaVersion { get; set; }
        [DataMember(Order = 2)] public List<LocalContentItem> Items { get; set; }
    }

    public sealed class LocalContentStore
    {
        private readonly JsonSettingsStore<LocalContentData> _store;
        public LocalContentStore(string path) { _store = new JsonSettingsStore<LocalContentData>(path); }

        public IReadOnlyList<LocalContentItem> Search(LocalContentKind kind, string query)
        {
            string text = (query ?? string.Empty).Trim();
            return Read().Items.Where(item => item.Kind == kind)
                .Where(item => text.Length == 0 || Contains(item.Title, text) || Contains(item.Tags, text) || Contains(item.Content, text))
                .OrderByDescending(item => item.UpdatedUtc).ToList();
        }

        public LocalContentItem Add(LocalContentKind kind, string title, string content, string tags)
        {
            Validate(title, content);
            var data = Read();
            var now = DateTime.UtcNow;
            var item = new LocalContentItem { Id = Guid.NewGuid().ToString("N"), Kind = kind, Title = title.Trim(), Content = content, Tags = tags ?? string.Empty, CreatedUtc = now, UpdatedUtc = now };
            data.Items.Add(item);
            Write(data);
            return item;
        }

        public void Update(string id, string title, string content, string tags)
        {
            Validate(title, content);
            var data = Read();
            var item = data.Items.SingleOrDefault(x => string.Equals(x.Id, id, StringComparison.Ordinal));
            if (item == null) throw new InvalidOperationException("要更新的条目不存在。");
            item.Title = title.Trim();
            item.Content = content;
            item.Tags = tags ?? string.Empty;
            item.UpdatedUtc = DateTime.UtcNow;
            Write(data);
        }

        public void Delete(string id)
        {
            var data = Read();
            int removed = data.Items.RemoveAll(item => string.Equals(item.Id, id, StringComparison.Ordinal));
            if (removed == 0) throw new InvalidOperationException("要删除的条目不存在。");
            Write(data);
        }

        private LocalContentData Read()
        {
            var data = _store.Load();
            if (data.Items == null) data.Items = new List<LocalContentItem>();
            return data;
        }

        private void Write(LocalContentData data) { _store.Save(data); }
        private static bool Contains(string value, string query) { return (value ?? string.Empty).IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0; }

        private static void Validate(string title, string content)
        {
            if (string.IsNullOrWhiteSpace(title) || title.Length > 200) throw new ArgumentException("标题应为 1 至 200 个字符。", "title");
            if (string.IsNullOrWhiteSpace(content) || content.Length > 500000) throw new ArgumentException("内容应为 1 至 500000 个字符。", "content");
        }
    }
}
