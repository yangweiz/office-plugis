using Suite.Core.Commands;

namespace Suite.Windowing
{
    public static class WindowIntegration
    {
        public static CommandResult GetNativeTabStatus()
        {
            return CommandResult.Blocked("单窗口原生标签、拖拽分组及多窗口合并需要独立的 Win32/COM 集成，并须完成 Office 版本、DPI 和多显示器验证。当前仅提供文档切换侧栏，未宣称原生标签可用。");
        }
    }
}
