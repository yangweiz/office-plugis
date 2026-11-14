using System;
using System.Threading;

namespace Suite.Core.Commands
{
    public sealed class CommandExecutor
    {
        public CommandResult Execute(ISuiteCommand command, CommandContext context, CancellationToken cancellationToken)
        {
            if (command == null) return CommandResult.Failed("未找到请求的功能。");
            if (context == null) return CommandResult.Failed("功能上下文不可用。");
            if (!command.CanExecute(context)) return CommandResult.Blocked("当前文档或选区不满足此功能的执行条件。");
            if (cancellationToken.IsCancellationRequested) return CommandResult.Cancelled("操作已取消。");
            try
            {
                return command.Execute(context, cancellationToken) ?? CommandResult.Failed("功能未返回执行结果。");
            }
            catch (OperationCanceledException)
            {
                return CommandResult.Cancelled("操作已取消；已完成的输出会保留并列入结果。");
            }
            catch (UnauthorizedAccessException ex)
            {
                return CommandResult.Failed("权限不足：" + ex.Message);
            }
            catch (Exception ex)
            {
                return CommandResult.Failed(command.DisplayName + "执行失败：" + ex.Message);
            }
        }
    }
}
