using System.Threading;

namespace Suite.Core.Commands
{
    public interface ISuiteCommand
    {
        string Id { get; }
        string DisplayName { get; }
        string Keywords { get; }
        SuiteHost Hosts { get; }
        CommandRisk Risk { get; }
        bool CanUndo { get; }
        bool CanExecute(CommandContext context);
        CommandResult Execute(CommandContext context, CancellationToken cancellationToken);
    }
}
