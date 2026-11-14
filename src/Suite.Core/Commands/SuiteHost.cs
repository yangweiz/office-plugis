using System;

namespace Suite.Core.Commands
{
    [Flags]
    public enum SuiteHost
    {
        None = 0,
        Word = 1,
        Excel = 2,
        PowerPoint = 4,
        All = Word | Excel | PowerPoint
    }
}
