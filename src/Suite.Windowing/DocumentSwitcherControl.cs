using System;
using System.Drawing;
using System.Windows.Forms;

namespace Suite.Windowing
{
    public sealed class DocumentSwitcherControl : UserControl
    {
        private readonly DocumentSessionService _service;
        private readonly ListBox _list = new ListBox();
        private bool _refreshing;

        public DocumentSwitcherControl(DocumentSessionService service)
        {
            _service = service ?? throw new ArgumentNullException("service");
            Dock = DockStyle.Fill;
            Font = new Font("Microsoft YaHei UI", 9F);
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 72, FlowDirection = FlowDirection.LeftToRight, WrapContents = true, Padding = new Padding(4) };
            _list.Dock = DockStyle.Fill;
            _list.DrawMode = DrawMode.OwnerDrawFixed;
            _list.ItemHeight = 28;
            _list.IntegralHeight = false;
            _list.DrawItem += OnDrawItem;
            _list.DoubleClick += delegate { Run(delegate { if (Selected != null) _service.Activate(Selected); }); };
            Controls.Add(_list);
            Controls.Add(buttons);
            AddButton(buttons, "激活", delegate { if (Selected != null) _service.Activate(Selected); });
            AddButton(buttons, "保存", delegate { if (Selected != null) _service.Save(Selected); });
            AddButton(buttons, "关闭", delegate { if (Selected != null) _service.Close(Selected, ConfirmClose); });
            AddButton(buttons, "全部保存", delegate { _service.SaveAll(); });
            AddButton(buttons, "全部关闭", delegate { _service.CloseAll(ConfirmClose); });
            AddButton(buttons, "颜色", delegate { ChooseColor(); });
            _service.Changed += delegate { RefreshList(); };
            RefreshList();
        }

        public DocumentSession Selected { get { return _list.SelectedItem as DocumentSession; } }

        private void AddButton(Control parent, string text, Action action)
        {
            var button = new Button { Text = text, Width = 80, Height = 28, Margin = new Padding(3) };
            button.Click += delegate { Run(action); };
            parent.Controls.Add(button);
        }

        private void RefreshList()
        {
            _refreshing = true;
            try
            {
                string selectedKey = Selected == null ? null : Selected.Key;
                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (DocumentSession session in _service.Sessions) _list.Items.Add(session);
                foreach (DocumentSession session in _list.Items)
                    if (session.Key == selectedKey) { _list.SelectedItem = session; break; }
                if (_list.SelectedIndex < 0 && _list.Items.Count > 0) _list.SelectedIndex = 0;
                _list.EndUpdate();
            }
            finally { _refreshing = false; }
        }

        private void OnDrawItem(object sender, DrawItemEventArgs e)
        {
            if (_refreshing || e.Index < 0 || e.Index >= _list.Items.Count) return;
            var session = (DocumentSession)_list.Items[e.Index];
            Color background = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? SystemColors.Highlight : session.Color;
            if (background == Color.Empty) background = _list.BackColor;
            using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, e.Bounds);
            Color foreground = (e.State & DrawItemState.Selected) == DrawItemState.Selected ? SystemColors.HighlightText : session.Color != Color.Empty && session.Color.GetBrightness() < 0.45f ? Color.White : _list.ForeColor;
            TextRenderer.DrawText(e.Graphics, session.ToString(), Font, new Rectangle(e.Bounds.X + 6, e.Bounds.Y, e.Bounds.Width - 6, e.Bounds.Height), foreground, TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
            e.DrawFocusRectangle();
        }

        private static CloseDecision ConfirmClose(DocumentSession session)
        {
            DialogResult result = MessageBox.Show("“" + session.DisplayName + "”尚未保存。\n是：保存并关闭；否：放弃更改并关闭；取消：保持打开。", "关闭文档", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) return CloseDecision.Save;
            if (result == DialogResult.No) return CloseDecision.Discard;
            return CloseDecision.Cancel;
        }

        private void ChooseColor()
        {
            if (Selected == null) return;
            using (var dialog = new ColorDialog { Color = Selected.Color == Color.Empty ? Color.White : Selected.Color, FullOpen = true })
            {
                if (dialog.ShowDialog() == DialogResult.OK) Run(delegate { _service.SetColor(Selected, dialog.Color); });
            }
        }

        private void Run(Action action)
        {
            try { action(); }
            catch (Exception ex) { MessageBox.Show(ex.Message, "文档管理", MessageBoxButtons.OK, MessageBoxIcon.Warning); }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) Font.Dispose();
            base.Dispose(disposing);
        }
    }
}
