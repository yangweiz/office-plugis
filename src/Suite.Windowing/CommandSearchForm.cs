using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Suite.Core.Commands;

namespace Suite.Windowing
{
    public sealed class CommandSearchForm : Form
    {
        private readonly CommandRegistry _registry;
        private readonly SuiteHost _host;
        private readonly TextBox _query;
        private readonly ListBox _results;
        private readonly Label _detail;
        private readonly Button _run;

        public CommandSearchForm(CommandRegistry registry, SuiteHost host, string initialQuery)
        {
            _registry = registry ?? throw new ArgumentNullException("registry");
            _host = host;
            Text = "功能搜索";
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            MaximizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(560, 390);
            Font = new Font("Microsoft YaHei UI", 9F);
            var label = new Label { Text = "输入功能名称、功能 ID 或关键词", Left = 16, Top = 16, Width = 520 };
            _query = new TextBox { Left = 16, Top = 42, Width = 528, Text = initialQuery ?? string.Empty };
            _results = new ListBox { Left = 16, Top = 74, Width = 528, Height = 230, IntegralHeight = false };
            _detail = new Label { Left = 16, Top = 312, Width = 360, Height = 56, AutoEllipsis = true };
            _run = new Button { Text = "执行", Left = 386, Top = 324, Width = 74, Height = 32, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "取消", Left = 470, Top = 324, Width = 74, Height = 32, DialogResult = DialogResult.Cancel };
            Controls.AddRange(new Control[] { label, _query, _results, _detail, _run, cancel });
            AcceptButton = _run;
            CancelButton = cancel;
            _query.TextChanged += delegate { RefreshResults(); };
            _results.SelectedIndexChanged += delegate { UpdateDetail(); };
            _results.DoubleClick += delegate { if (SelectedCommand != null) DialogResult = DialogResult.OK; };
            _run.Click += delegate { if (SelectedCommand == null) DialogResult = DialogResult.None; };
            RefreshResults();
        }

        public ISuiteCommand SelectedCommand
        {
            get
            {
                CommandItem item = _results.SelectedItem as CommandItem;
                return item == null ? null : item.Command;
            }
        }

        private void RefreshResults()
        {
            IReadOnlyList<ISuiteCommand> commands = _registry.Search(_host, _query.Text);
            _results.BeginUpdate();
            _results.Items.Clear();
            foreach (ISuiteCommand command in commands) _results.Items.Add(new CommandItem(command));
            _results.EndUpdate();
            if (_results.Items.Count > 0) _results.SelectedIndex = 0;
            UpdateDetail();
        }

        private void UpdateDetail()
        {
            ISuiteCommand command = SelectedCommand;
            _detail.Text = command == null ? "没有匹配功能。" : command.Id + " · " + (command.CanUndo ? "支持宿主撤销" : "修改前备份或明确确认");
            _run.Enabled = command != null;
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _query.Focus();
            _query.SelectAll();
        }

        private sealed class CommandItem
        {
            public CommandItem(ISuiteCommand command) { Command = command; }
            public ISuiteCommand Command { get; private set; }
            public override string ToString() { return Command.DisplayName; }
        }
    }
}
