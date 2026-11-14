namespace Suite.Word
{
    [Microsoft.VisualStudio.Tools.Applications.Runtime.StartupObjectAttribute(0)]
    [System.Security.Permissions.PermissionSetAttribute(System.Security.Permissions.SecurityAction.Demand, Name = "FullTrust")]
    public sealed partial class ThisAddIn : Microsoft.Office.Tools.AddInBase
    {
        internal Microsoft.Office.Interop.Word.Application Application;

        public ThisAddIn(Microsoft.Office.Tools.Factory factory, System.IServiceProvider serviceProvider)
            : base(factory, serviceProvider, "AddIn", "ThisAddIn")
        {
        }

        protected override void Initialize()
        {
            base.Initialize();
            Application = (Microsoft.Office.Interop.Word.Application)HostItem;
        }

        protected override void FinishInitialization()
        {
            InternalStartup();
            base.FinishInitialization();
        }

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }
    }
}
