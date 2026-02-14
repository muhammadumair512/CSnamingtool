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
            SendMessage(textBox.Handle, EM_SETCUEBANNER, 0, placeholder);
        }

        private PowerPoint.Application pptApp;
        
        // UI Controls
        private ListView lstShapes;
        private CheckBox chkFocusMode;
        private CheckBox chkInverse;
        private Button btnRefresh;
        private TextBox txtSearch;
        
        // Alignment Buttons
        private Button btnAlignLeft, btnAlignRight, btnAlignTop, btnAlignBottom, btnAlignCenter, btnAlignMiddle;
        private Button btnDistributeH, btnDistributeV;
        private Button btnGroup, btnUngroup;
        
        // State
        private Dictionary<int, Office.MsoTriState> originalVisibility = new Dictionary<int, Office.MsoTriState>();
        private bool isProcessingSelection = false;

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
            this.Size = new Size(550, 850);
            this.MinimumSize = new Size(500, 600);
            this.ShowIcon = false;

            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 4;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 40f)); // Search
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // List
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 110f)); // Tools
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 60f));  // Options
            this.Controls.Add(mainLayout);

            // 1. Search Bar
            Panel searchPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            txtSearch = new TextBox { Dock = DockStyle.Fill };
            SetPlaceholder(txtSearch, "Search shapes by name...");
            txtSearch.TextChanged += (s, e) => FilterShapes();
            searchPanel.Controls.Add(txtSearch);
            mainLayout.Controls.Add(searchPanel, 0, 0);

            // 2. List View
            lstShapes = new ListView();
            lstShapes.Dock = DockStyle.Fill;
            lstShapes.View = View.Details;
            lstShapes.FullRowSelect = true;
            lstShapes.MultiSelect = true;
            lstShapes.AllowDrop = true;
            lstShapes.GridLines = true;
            lstShapes.LabelEdit = true; 
            
            lstShapes.Columns.Add("Z", 40);
            lstShapes.Columns.Add("Shape Name", 180);
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

            // 3. Tools Group
            GroupBox grpTools = new GroupBox { Text = "Industrial Tools", Dock = DockStyle.Fill, Margin = new Padding(5) };
            FlowLayoutPanel flowTools = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight };
            grpTools.Controls.Add(flowTools);
            mainLayout.Controls.Add(grpTools, 0, 2);

            btnAlignLeft = CreateToolButton("Left", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignLefts));
            btnAlignCenter = CreateToolButton("Center", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignCenters));
            btnAlignRight = CreateToolButton("Right", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignRights));
            btnAlignTop = CreateToolButton("Top", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignTops));
            btnAlignMiddle = CreateToolButton("Middle", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignMiddles));
            btnAlignBottom = CreateToolButton("Bottom", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignBottoms));
            btnDistributeH = CreateToolButton("Dist H", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeHorizontally));
            btnDistributeV = CreateToolButton("Dist V", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeVertically));
            btnGroup = CreateToolButton("Group", (s, e) => GroupSelected());
            btnUngroup = CreateToolButton("Ungroup", (s, e) => UngroupSelected());

            flowTools.Controls.AddRange(new Control[] { btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignTop, btnAlignMiddle, btnAlignBottom, btnDistributeH, btnDistributeV, btnGroup, btnUngroup });

            // 4. Options Panel
            FlowLayoutPanel flowOptions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(5) };
            mainLayout.Controls.Add(flowOptions, 0, 3);

            chkFocusMode = new CheckBox { Text = "Focus Mode", AutoSize = true, Margin = new Padding(5) };
            chkFocusMode.CheckedChanged += (s, e) => ApplyVisibility();
            
            chkInverse = new CheckBox { Text = "Inverse", AutoSize = true, Margin = new Padding(5) };
            chkInverse.CheckedChanged += (s, e) => ApplyVisibility();

            btnRefresh = new Button { Text = "Refresh", AutoSize = true, FlatStyle = FlatStyle.System };
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
            LoadShapes();
        }

        private void ShapeManagerForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            RestoreOriginalVisibility();
        }

        public void SyncSelectionFromPowerPoint(PowerPoint.Selection sel)
        {
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => SyncSelectionFromPowerPoint(sel)));
                return;
            }

            if (isProcessingSelection) return;
            isProcessingSelection = true;

            try
            {
                lstShapes.SelectedIndexChanged -= LstShapes_SelectedIndexChanged;
                lstShapes.SelectedItems.Clear();
                
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    HashSet<int> selectedIds = new HashSet<int>();
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                    {
                        selectedIds.Add(shape.Id);
                    }

                    lstShapes.BeginUpdate();
                    foreach (ListViewItem item in lstShapes.Items)
                    {
                        if (item.Tag is PowerPoint.Shape shape && selectedIds.Contains(shape.Id))
                        {
                            item.Selected = true;
                        }
                    }
                    if (lstShapes.SelectedItems.Count > 0)
                        lstShapes.SelectedItems[0].EnsureVisible();
                    lstShapes.EndUpdate();
                }
            }
            catch { }
            finally
            {
                lstShapes.SelectedIndexChanged += LstShapes_SelectedIndexChanged;
                isProcessingSelection = false;
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

                for (int i = slide.Shapes.Count; i >= 1; i--)
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
            }
            catch { }
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
                    {
                        range.Distribute(distCmd, Office.MsoTriState.msoFalse);
                    }
                    else
                    {
                        MessageBox.Show("Distribution requires at least 2 shapes.");
                    }
                }
            }
            catch (Exception ex) { MessageBox.Show("Distribute Error: " + ex.Message); }
        }

        private void GroupSelected()
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    pptApp.ActiveWindow.Selection.ShapeRange.Group();
                    LoadShapes();
                }
            }
            catch (Exception ex) { MessageBox.Show("Group Error: " + ex.Message); }
        }

        private void UngroupSelected()
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    pptApp.ActiveWindow.Selection.ShapeRange.Ungroup();
                    LoadShapes();
                }
            }
            catch (Exception ex) { MessageBox.Show("Ungroup Error: " + ex.Message); }
        }

        private void LstShapes_AfterLabelEdit(object sender, LabelEditEventArgs e)
        {
            if (e.Label == null) return;
            try
            {
                ListViewItem item = lstShapes.Items[e.Item];
                if (item.Tag is PowerPoint.Shape shape)
                {
                    shape.Name = e.Label;
                }
            }
            catch (Exception ex) 
            { 
                MessageBox.Show("Rename Error: " + ex.Message);
                e.CancelEdit = true;
            }
        }

        private PowerPoint.Slide GetActiveSlide()
        {
            try { return (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide; } catch { return null; }
        }

        private void LstShapes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isProcessingSelection) return;
            isProcessingSelection = true;
            try
            {
                UpdatePPSelection();
                ApplyVisibility();
            }
            finally { isProcessingSelection = false; }
        }

        private void UpdatePPSelection()
        {
            try
            {
                pptApp.ActiveWindow.Selection.Unselect();
                foreach (ListViewItem item in lstShapes.SelectedItems)
                {
                    if (item.Tag is PowerPoint.Shape shape)
                        shape.Select(Office.MsoTriState.msoFalse);
                }
            }
            catch { }
        }

        private void ApplyVisibility()
        {
            if (!chkFocusMode.Checked)
            {
                RestoreOriginalVisibility();
                return;
            }

            var slide = GetActiveSlide();
            if (slide == null) return;

            HashSet<int> selectedIds = new HashSet<int>();
            foreach (ListViewItem item in lstShapes.SelectedItems)
            {
                if (item.Tag is PowerPoint.Shape shape)
                {
                    selectedIds.Add(shape.Id);
                }
            }

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (!originalVisibility.ContainsKey(shape.Id))
                {
                    originalVisibility[shape.Id] = shape.Visible;
                }

                bool shouldBeVisible;
                if (chkInverse.Checked)
                {
                    shouldBeVisible = !selectedIds.Contains(shape.Id);
                }
                else
                {
                    shouldBeVisible = selectedIds.Contains(shape.Id);
                }

                shape.Visible = shouldBeVisible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }
        }

        private void RestoreOriginalVisibility()
        {
            if (originalVisibility.Count == 0) return;

            var slide = GetActiveSlide();
            if (slide == null) return;

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if (originalVisibility.TryGetValue(shape.Id, out Office.MsoTriState state))
                {
                    shape.Visible = state;
                }
            }
            originalVisibility.Clear();
        }

        #region Drag & Drop for Z-Order

        private void LstShapes_ItemDrag(object sender, ItemDragEventArgs e)
        {
            lstShapes.DoDragDrop(e.Item, DragDropEffects.Move);
        }

        private void LstShapes_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(ListViewItem))) 
                e.Effect = DragDropEffects.Move;
        }

        private void LstShapes_DragLeave(object sender, EventArgs e)
        {
            lstShapes.InsertionMark.Index = -1;
        }

        private void LstShapes_DragOver(object sender, DragEventArgs e)
        {
            Point cp = lstShapes.PointToClient(new Point(e.X, e.Y));
            ListViewItem targetItem = lstShapes.GetItemAt(cp.X, cp.Y);

            if (targetItem != null)
            {
                int targetIndex = targetItem.Index;
                Rectangle itemBounds = targetItem.GetBounds(ItemBoundsPortion.Entire);
                
                if (cp.Y > itemBounds.Top + (itemBounds.Height / 2))
                {
                    lstShapes.InsertionMark.AppearsAfterItem = true;
                    lstShapes.InsertionMark.Index = targetIndex;
                }
                else
                {
                    lstShapes.InsertionMark.AppearsAfterItem = false;
                    lstShapes.InsertionMark.Index = targetIndex;
                }
                
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
                lstShapes.InsertionMark.Index = -1;
            }
        }

        private void LstShapes_DragDrop(object sender, DragEventArgs e)
        {
            int insertionIndex = lstShapes.InsertionMark.Index;
            bool after = lstShapes.InsertionMark.AppearsAfterItem;
            lstShapes.InsertionMark.Index = -1; 

            if (insertionIndex == -1) return;

            try
            {
                ListViewItem draggedItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));
                if (draggedItem == null) return;
                
                PowerPoint.Shape shapeToMove = draggedItem.Tag as PowerPoint.Shape;
                ListViewItem targetItem = lstShapes.Items[insertionIndex];
                PowerPoint.Shape targetShape = targetItem.Tag as PowerPoint.Shape;

                if (shapeToMove != null && targetShape != null)
                {
                    int currentPos = shapeToMove.ZOrderPosition;
                    int targetPos = targetShape.ZOrderPosition;
                    int finalTargetPos = after ? targetPos - 1 : targetPos;
                    
                    int moves = 0;
                    if (currentPos < finalTargetPos)
                    {
                        moves = finalTargetPos - currentPos;
                        for (int i = 0; i < moves; i++) shapeToMove.ZOrder(Office.MsoZOrderCmd.msoBringForward);
                    }
                    else if (currentPos > finalTargetPos)
                    {
                        moves = currentPos - finalTargetPos;
                        for (int i = 0; i < moves; i++) shapeToMove.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                    }
                    
                    LoadShapes();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reordering shapes: " + ex.Message);
            }
        }

        #endregion
    }
}
