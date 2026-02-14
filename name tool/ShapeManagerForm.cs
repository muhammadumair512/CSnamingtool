using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace name_tool
{
    public class ShapeManagerForm : Form
    {
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, int wParam, [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPWStr)] string lParam);
        private const int EM_SETCUEBANNER = 0x1501;

        private void SetPlaceholder(TextBox textBox, string placeholder)
        {
            if (textBox.IsHandleCreated)
                SendMessage(textBox.Handle, EM_SETCUEBANNER, 0, placeholder);
        }

        private PowerPoint.Application pptApp;
        
        // UI Controls
        private ListView lstShapes;
        private CheckBox chkFocusMode;
        private CheckBox chkInverse;
        private Button btnRefresh;
        private TextBox txtSearch;
        
        // Buttons
        private Button btnAlignLeft, btnAlignRight, btnAlignTop, btnAlignBottom, btnAlignCenter, btnAlignMiddle;
        private Button btnDistributeH, btnDistributeV;
        private Button btnGroup, btnUngroup, btnDelete, btnSelectAll;
        private Button btnMatchWidth, btnMatchHeight, btnSwap, btnSelectSameType;
        private Button btnHideAll, btnShowAll;
        
        // State
        private Dictionary<int, Office.MsoTriState> originalVisibility = new Dictionary<int, Office.MsoTriState>();
        private bool isInternalChange = false;

        public ShapeManagerForm(PowerPoint.Application app)
        {
            this.pptApp = app;
            InitializeComponent();
            this.TopMost = true;
            this.Load += ShapeManagerForm_Load;
            this.FormClosing += ShapeManagerForm_FormClosing;
        }

        private void InitializeComponent()
        {
            this.Text = "Advanced Shape Manager Pro";
            this.Size = new Size(600, 950);
            this.MinimumSize = new Size(550, 750);
            this.ShowIcon = false;

            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 5;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f)); // Search
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // List
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110f)); // Layout
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 150f)); // Advanced Efficiency
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60f));  // Options
            this.Controls.Add(mainLayout);

            Panel searchPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            txtSearch = new TextBox { Dock = DockStyle.Fill };
            txtSearch.TextChanged += (s, e) => FilterShapes();
            searchPanel.Controls.Add(txtSearch);
            mainLayout.Controls.Add(searchPanel, 0, 0);

            lstShapes = new ListView();
            lstShapes.Dock = DockStyle.Fill;
            lstShapes.View = View.Details;
            lstShapes.FullRowSelect = true;
            lstShapes.MultiSelect = true;
            lstShapes.AllowDrop = true;
            lstShapes.GridLines = true;
            lstShapes.LabelEdit = true; 
            
            lstShapes.Columns.Add("Z", 40);
            lstShapes.Columns.Add("Shape Name", 220);
            lstShapes.Columns.Add("Type", 90);
            lstShapes.Columns.Add("W", 50);
            lstShapes.Columns.Add("H", 50);
            
            lstShapes.SelectedIndexChanged += LstShapes_SelectedIndexChanged;
            lstShapes.AfterLabelEdit += LstShapes_AfterLabelEdit;
            lstShapes.ItemDrag += LstShapes_ItemDrag;
            lstShapes.DragEnter += LstShapes_DragEnter;
            lstShapes.DragDrop += LstShapes_DragDrop;
            lstShapes.DragOver += LstShapes_DragOver;
            lstShapes.DragLeave += LstShapes_DragLeave;
            
            mainLayout.Controls.Add(lstShapes, 0, 1);

            GroupBox grpAlign = new GroupBox { Text = "Layout & Distribution", Dock = DockStyle.Fill, Margin = new Padding(5) };
            FlowLayoutPanel flowAlign = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            grpAlign.Controls.Add(flowAlign);
            mainLayout.Controls.Add(grpAlign, 0, 2);

            flowAlign.Controls.Add(CreateToolButton("Left", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignLefts)));
            flowAlign.Controls.Add(CreateToolButton("Center", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignCenters)));
            flowAlign.Controls.Add(CreateToolButton("Right", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignRights)));
            flowAlign.Controls.Add(CreateToolButton("Top", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignTops)));
            flowAlign.Controls.Add(CreateToolButton("Middle", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignMiddles)));
            flowAlign.Controls.Add(CreateToolButton("Bottom", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignBottoms)));
            flowAlign.Controls.Add(CreateToolButton("Dist H", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeHorizontally)));
            flowAlign.Controls.Add(CreateToolButton("Dist V", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeVertically)));

            GroupBox grpAdvanced = new GroupBox { Text = "Advanced Efficiency Tools", Dock = DockStyle.Fill, Margin = new Padding(5) };
            FlowLayoutPanel flowAdvanced = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            grpAdvanced.Controls.Add(flowAdvanced);
            mainLayout.Controls.Add(grpAdvanced, 0, 3);

            flowAdvanced.Controls.Add(CreateToolButton("Match W", (s, e) => MatchSize(true, false)));
            flowAdvanced.Controls.Add(CreateToolButton("Match H", (s, e) => MatchSize(false, true)));
            flowAdvanced.Controls.Add(CreateToolButton("Swap Pos", (s, e) => SwapShapes()));
            flowAdvanced.Controls.Add(CreateToolButton("Same Type", (s, e) => SelectSameType()));
            flowAdvanced.Controls.Add(CreateToolButton("Select All", (s, e) => SelectAllShapes()));
            flowAdvanced.Controls.Add(CreateToolButton("Group", (s, e) => GroupSelected()));
            flowAdvanced.Controls.Add(CreateToolButton("Ungroup", (s, e) => UngroupSelected()));
            btnDelete = CreateToolButton("Delete", (s, e) => DeleteSelected());
            btnDelete.BackColor = Color.MistyRose;
            flowAdvanced.Controls.Add(btnDelete);
            flowAdvanced.Controls.Add(CreateToolButton("Hide All", (s, e) => ToggleAllVisibility(false)));
            flowAdvanced.Controls.Add(CreateToolButton("Show All", (s, e) => ToggleAllVisibility(true)));

            FlowLayoutPanel flowOptions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            mainLayout.Controls.Add(flowOptions, 0, 4);

            chkFocusMode = new CheckBox { Text = "Focus Mode", AutoSize = true, Margin = new Padding(5) };
            chkFocusMode.CheckedChanged += (s, e) => ApplyVisibility();
            
            chkInverse = new CheckBox { Text = "Inverse", AutoSize = true, Margin = new Padding(5) };
            chkInverse.CheckedChanged += (s, e) => ApplyVisibility();

            btnRefresh = new Button { Text = "Refresh List", AutoSize = true, FlatStyle = FlatStyle.System };
            btnRefresh.Click += (s, e) => LoadShapes();

            flowOptions.Controls.AddRange(new Control[] { chkFocusMode, chkInverse, btnRefresh });
        }

        private Button CreateToolButton(string text, EventHandler onClick)
        {
            Button btn = new Button { Text = text, Width = 65, Height = 30, Margin = new Padding(2), FlatStyle = FlatStyle.Flat };
            btn.Click += onClick;
            return btn;
        }

        private void ShapeManagerForm_Load(object sender, EventArgs e)
        {
            SetPlaceholder(txtSearch, "Search shapes by name...");
            LoadShapes();
        }

        private void ShapeManagerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            RestoreOriginalVisibility();
        }

        public void SyncSelectionFromPowerPoint(PowerPoint.Selection sel)
        {
            if (isInternalChange) return;
            
            if (this.InvokeRequired)
            {
                this.Invoke(new Action(() => SyncSelectionFromPowerPoint(sel)));
                return;
            }

            try
            {
                isInternalChange = true;
                lstShapes.SelectedIndexChanged -= LstShapes_SelectedIndexChanged;
                
                HashSet<int> selectedIds = new HashSet<int>();
                if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        try { selectedIds.Add(shape.Id); } catch { }
                    }
                }

                lstShapes.BeginUpdate();
                foreach (ListViewItem item in lstShapes.Items)
                {
                    if (item.Tag is PowerPoint.Shape shape)
                    {
                        try
                        {
                            bool shouldBeSelected = selectedIds.Contains(shape.Id);
                            if (item.Selected != shouldBeSelected)
                                item.Selected = shouldBeSelected;
                        }
                        catch { }
                    }
                }
                
                if (lstShapes.SelectedItems.Count > 0)
                    lstShapes.SelectedItems[0].EnsureVisible();
                
                lstShapes.EndUpdate();
            }
            catch { }
            finally
            {
                lstShapes.SelectedIndexChanged += LstShapes_SelectedIndexChanged;
                isInternalChange = false;
            }
        }

        private void FilterShapes()
        {
            string filter = txtSearch.Text.ToLower();
            lstShapes.BeginUpdate();
            foreach (ListViewItem item in lstShapes.Items)
            {
                if (string.IsNullOrEmpty(filter)) item.BackColor = SystemColors.Window;
                else if (item.SubItems[1].Text.ToLower().Contains(filter)) item.BackColor = Color.LightYellow;
                else item.BackColor = SystemColors.Window;
            }
            lstShapes.EndUpdate();
        }

        private void LoadShapes()
        {
            lstShapes.BeginUpdate();
            lstShapes.Items.Clear();
            try
            {
                var slide = GetActiveSlide();
                if (slide == null) return;

                int count = 0;
                try { count = slide.Shapes.Count; } catch { }

                for (int i = count; i >= 1; i--)
                {
                    try
                    {
                        PowerPoint.Shape shape = slide.Shapes[i];
                        ListViewItem item = new ListViewItem(i.ToString());
                        item.SubItems.Add(shape.Name);
                        item.SubItems.Add(shape.Type.ToString().Replace("mso", ""));
                        item.SubItems.Add(Math.Round(shape.Width, 1).ToString());
                        item.SubItems.Add(Math.Round(shape.Height, 1).ToString());
                        item.Tag = shape;
                        lstShapes.Items.Add(item);
                    }
                    catch { continue; }
                }
            }
            catch (Exception ex) { MessageBox.Show("Load Error: " + ex.Message); }
            finally { lstShapes.EndUpdate(); }
        }

        private void AlignSelected(Office.MsoAlignCmd alignCmd)
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange range = pptApp.ActiveWindow.Selection.ShapeRange;
                    if (range.Count > 0)
                    {
                        Office.MsoTriState relativeToSlide = (range.Count == 1) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                        range.Align(alignCmd, relativeToSlide);
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Align Error: " + ex.Message); }
        }

        private void DistributeSelected(Office.MsoDistributeCmd distCmd)
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.ShapeRange range = pptApp.ActiveWindow.Selection.ShapeRange;
                    if (range.Count >= 2)
                        range.Distribute(distCmd, Office.MsoTriState.msoFalse);
                    else
                        MessageBox.Show("Select at least 2 shapes.");
                }
            }
            catch (Exception ex) { MessageBox.Show("Distribute Error: " + ex.Message); }
        }

        private void MatchSize(bool width, bool height)
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var range = pptApp.ActiveWindow.Selection.ShapeRange;
                    if (range.Count < 2) return;
                    float refW = range[1].Width, refH = range[1].Height;
                    for (int i = 2; i <= range.Count; i++)
                    {
                        if (width) range[i].Width = refW;
                        if (height) range[i].Height = refH;
                    }
                    LoadShapes();
                }
            }
            catch (Exception ex) { MessageBox.Show("Match Size Error: " + ex.Message); }
        }

        private void SwapShapes()
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var range = pptApp.ActiveWindow.Selection.ShapeRange;
                    if (range.Count != 2) { MessageBox.Show("Select 2 shapes."); return; }
                    float tL = range[1].Left, tT = range[1].Top;
                    range[1].Left = range[2].Left; range[1].Top = range[2].Top;
                    range[2].Left = tL; range[2].Top = tT;
                }
            }
            catch (Exception ex) { MessageBox.Show("Swap Error: " + ex.Message); }
        }

        private void SelectSameType()
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var type = pptApp.ActiveWindow.Selection.ShapeRange[1].Type;
                    var slide = GetActiveSlide();
                    if (slide == null) return;
                    List<string> names = new List<string>();
                    foreach (PowerPoint.Shape s in slide.Shapes) if (s.Type == type) names.Add(s.Name);
                    if (names.Count > 0) { isInternalChange = true; slide.Shapes.Range(names.ToArray()).Select(); isInternalChange = false; SyncSelectionFromPowerPoint(pptApp.ActiveWindow.Selection); }
                }
            }
            catch (Exception ex) { MessageBox.Show("Select Type Error: " + ex.Message); }
        }

        private void SelectAllShapes()
        {
            try
            {
                var slide = GetActiveSlide();
                if (slide == null || slide.Shapes.Count == 0) return;
                isInternalChange = true;
                slide.Shapes.SelectAll();
                isInternalChange = false;
                SyncSelectionFromPowerPoint(pptApp.ActiveWindow.Selection);
            }
            catch { }
        }

        private void DeleteSelected()
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (MessageBox.Show("Delete selected items?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        isInternalChange = true;
                        pptApp.ActiveWindow.Selection.Delete();
                        isInternalChange = false;
                        LoadShapes();
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Delete Error: " + ex.Message); }
        }

        private void ToggleAllVisibility(bool visible)
        {
            try
            {
                var slide = GetActiveSlide();
                if (slide == null) return;
                foreach (PowerPoint.Shape s in slide.Shapes) try { s.Visible = visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
                LoadShapes();
            }
            catch { }
        }

        private void GroupSelected() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { pptApp.ActiveWindow.Selection.ShapeRange.Group(); LoadShapes(); } } catch { } }
        private void UngroupSelected() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { pptApp.ActiveWindow.Selection.ShapeRange.Ungroup(); LoadShapes(); } } catch { } }

        private void LstShapes_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            if (e.Label == null) return;
            try { ListViewItem item = lstShapes.Items[e.Item]; if (item.Tag is PowerPoint.Shape s) s.Name = e.Label; } catch (Exception ex) { MessageBox.Show("Rename Error: " + ex.Message); e.CancelEdit = true; }
        }

        private PowerPoint.Slide GetActiveSlide()
        {
            try
            {
                if (pptApp.ActiveWindow == null) return null;
                try { return (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide; } catch { }
                try { if (pptApp.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionNone) return (PowerPoint.Slide)pptApp.ActiveWindow.Selection.SlideRange[1]; } catch { }
            }
            catch { }
            return null;
        }

        private void LstShapes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isInternalChange) return;
            isInternalChange = true;
            try { UpdatePPSelection(); ApplyVisibility(); } finally { isInternalChange = false; }
        }

        private void UpdatePPSelection()
        {
            try
            {
                if (lstShapes.SelectedItems.Count == 0) { pptApp.ActiveWindow.Selection.Unselect(); return; }
                List<string> names = new List<string>();
                foreach (ListViewItem item in lstShapes.SelectedItems) if (item.Tag is PowerPoint.Shape s) names.Add(s.Name);
                var slide = GetActiveSlide();
                if (slide != null && names.Count > 0) slide.Shapes.Range(names.ToArray()).Select();
            }
            catch { }
        }

        private void ApplyVisibility()
        {
            if (!chkFocusMode.Checked) { RestoreOriginalVisibility(); return; }
            var slide = GetActiveSlide();
            if (slide == null) return;
            HashSet<int> selectedIds = new HashSet<int>();
            foreach (ListViewItem item in lstShapes.SelectedItems) if (item.Tag is PowerPoint.Shape s) selectedIds.Add(s.Id);
            foreach (PowerPoint.Shape s in slide.Shapes)
            {
                try {
                    if (!originalVisibility.ContainsKey(s.Id)) originalVisibility[s.Id] = s.Visible;
                    bool vis = chkInverse.Checked ? !selectedIds.Contains(s.Id) : selectedIds.Contains(s.Id);
                    s.Visible = vis ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                } catch { }
            }
        }

        private void RestoreOriginalVisibility()
        {
            if (originalVisibility.Count == 0) return;
            var slide = GetActiveSlide();
            if (slide == null) return;
            foreach (PowerPoint.Shape s in slide.Shapes) try { if (originalVisibility.TryGetValue(s.Id, out Office.MsoTriState state)) s.Visible = state; } catch { }
            originalVisibility.Clear();
        }

        #region Drag & Drop
        private void LstShapes_ItemDrag(object sender, ItemDragEventArgs e) { lstShapes.DoDragDrop(e.Item, DragDropEffects.Move); }
        private void LstShapes_DragEnter(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(typeof(ListViewItem))) e.Effect = DragDropEffects.Move; }
        private void LstShapes_DragLeave(object sender, EventArgs e) { lstShapes.InsertionMark.Index = -1; }
        private void LstShapes_DragOver(object sender, DragEventArgs e)
        {
            Point cp = lstShapes.PointToClient(new Point(e.X, e.Y));
            ListViewItem targetItem = lstShapes.GetItemAt(cp.X, cp.Y);
            if (targetItem != null)
            {
                int targetIndex = targetItem.Index;
                Rectangle itemBounds = targetItem.GetBounds(ItemBoundsPortion.Entire);
                lstShapes.InsertionMark.AppearsAfterItem = (cp.Y > itemBounds.Top + (itemBounds.Height / 2));
                lstShapes.InsertionMark.Index = targetIndex;
                e.Effect = DragDropEffects.Move;
            }
            else { e.Effect = DragDropEffects.None; lstShapes.InsertionMark.Index = -1; }
        }

        private void LstShapes_DragDrop(object sender, DragEventArgs e)
        {
            int insertionIndex = lstShapes.InsertionMark.Index;
            bool after = lstShapes.InsertionMark.AppearsAfterItem;
            lstShapes.InsertionMark.Index = -1; 
            if (insertionIndex == -1) return;
            try {
                ListViewItem draggedItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));
                PowerPoint.Shape sMove = draggedItem?.Tag as PowerPoint.Shape;
                PowerPoint.Shape sTarget = lstShapes.Items[insertionIndex].Tag as PowerPoint.Shape;
                if (sMove != null && sTarget != null) {
                    int cPos = sMove.ZOrderPosition, tPos = sTarget.ZOrderPosition;
                    int fPos = after ? tPos - 1 : tPos;
                    if (cPos < fPos) for (int i = 0; i < fPos - cPos; i++) sMove.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                    else if (cPos > fPos) for (int i = 0; i < cPos - fPos; i++) sMove.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    LoadShapes();
                }
            } catch { }
        }
        #endregion
    }
}
