using System;
using System.Collections.Generic;

namespace Suite.Core.Commands
{
    public enum CommandRisk
    {
        ReadOnly,
        ModifyDocument,
        ModifyFiles,
        ExternalService
    }

    public enum CommandStatus
    {
        Succeeded,
        Partial,
        Failed,
        Cancelled,
        Blocked
    }

    public sealed class CommandContext
    {
        public CommandContext(SuiteHost host, object application, object activeDocument, IDictionary<string, string> parameters, IProgress<string> progress)
        {
            Host = host;
            Application = application;
            ActiveDocument = activeDocument;
            Parameters = parameters ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            Progress = progress;
        }

        public SuiteHost Host { get; private set; }
        public object Application { get; private set; }
        public object ActiveDocument { get; private set; }
        public IDictionary<string, string> Parameters { get; private set; }
        public IProgress<string> Progress { get; private set; }

        public void Report(string message)
        {
            if (Progress != null) Progress.Report(message);
        }
    }

    public sealed class CommandResult
    {
        private CommandResult(CommandStatus status, string message, string outputPath)
        {
            Status = status;
            Message = message ?? string.Empty;
            OutputPath = outputPath;
        }

        public CommandStatus Status { get; private set; }
        public string Message { get; private set; }
        public string OutputPath { get; private set; }
        public static CommandResult Success(string message, string outputPath = null) { return new CommandResult(CommandStatus.Succeeded, message, outputPath); }
        public static CommandResult Partial(string message) { return new CommandResult(CommandStatus.Partial, message, null); }
        public static CommandResult Failed(string message) { return new CommandResult(CommandStatus.Failed, message, null); }
        public static CommandResult Cancelled(string message) { return new CommandResult(CommandStatus.Cancelled, message, null); }
        public static CommandResult Blocked(string message) { return new CommandResult(CommandStatus.Blocked, message, null); }
    }
}
