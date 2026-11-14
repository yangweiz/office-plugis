using System;
using System.Collections.Generic;
using System.Linq;

namespace Suite.Core.Commands
{
    public sealed class CommandRegistry
    {
        private readonly Dictionary<string, ISuiteCommand> _commands = new Dictionary<string, ISuiteCommand>(StringComparer.OrdinalIgnoreCase);

        public void Register(ISuiteCommand command)
        {
            if (command == null) throw new ArgumentNullException("command");
            if (_commands.ContainsKey(command.Id)) throw new InvalidOperationException("重复的功能 ID：" + command.Id);
            _commands.Add(command.Id, command);
        }

        public ISuiteCommand Find(string id)
        {
            ISuiteCommand command;
            return id != null && _commands.TryGetValue(id, out command) ? command : null;
        }

        public IReadOnlyList<ISuiteCommand> Search(SuiteHost host, string query)
        {
            string text = (query ?? string.Empty).Trim();
            return _commands.Values
                .Where(c => (c.Hosts & host) == host)
                .Where(c => text.Length == 0 || c.DisplayName.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0 || c.Keywords.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0 || c.Id.IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                .OrderBy(c => c.DisplayName, StringComparer.CurrentCulture)
                .ToList();
        }
    }
}
