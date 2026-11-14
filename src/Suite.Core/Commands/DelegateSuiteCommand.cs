using System;
using System.Threading;

namespace Suite.Core.Commands
{
    public sealed class DelegateSuiteCommand : ISuiteCommand
    {
        private readonly Func<CommandContext, bool> _canExecute;
        private readonly Func<CommandContext, CancellationToken, CommandResult> _execute;

        public DelegateSuiteCommand(string id, string displayName, string keywords, SuiteHost hosts, CommandRisk risk, bool canUndo, Func<CommandContext, bool> canExecute, Func<CommandContext, CancellationToken, CommandResult> execute)
        {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("功能 ID 不能为空。", "id");
            if (string.IsNullOrWhiteSpace(displayName)) throw new ArgumentException("功能名称不能为空。", "displayName");
            if (execute == null) throw new ArgumentNullException("execute");
            Id = id;
            DisplayName = displayName;
            Keywords = keywords ?? string.Empty;
            Hosts = hosts;
            Risk = risk;
            CanUndo = canUndo;
            _canExecute = canExecute;
            _execute = execute;
        }

        public string Id { get; private set; }
        public string DisplayName { get; private set; }
        public string Keywords { get; private set; }
        public SuiteHost Hosts { get; private set; }
        public CommandRisk Risk { get; private set; }
        public bool CanUndo { get; private set; }
        public bool CanExecute(CommandContext context) { return context != null && (Hosts & context.Host) == context.Host && (_canExecute == null || _canExecute(context)); }
        public CommandResult Execute(CommandContext context, CancellationToken cancellationToken) { return _execute(context, cancellationToken); }
    }
}
