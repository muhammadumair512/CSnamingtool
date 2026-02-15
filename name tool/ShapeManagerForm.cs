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

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool BringWindowToTop(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr processId);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

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
        private Button btnMatchBoth, btnStackH, btnStackV, btnRotate90;
        private Button btnHideAll, btnShowAll;
        private Button btnToFront, btnToBack, btnForward, btnBackward, btnCenterH, btnCenterV;
        private Button btnFreeform, btnRect, btnLine;

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
            this.Size = new Size(450, 980);
            this.MinimumSize = new Size(420, 800);
            this.ShowIcon = false;

            TableLayoutPanel mainLayout = new TableLayoutPanel();
            mainLayout.Dock = DockStyle.Fill;
            mainLayout.RowCount = 6;
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 35f)); // Search
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // List
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 125f)); // Layout
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 85f));  // Drawing Tools
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 185f)); // Efficiency
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50f));  // Options
            this.Controls.Add(mainLayout);

            Panel searchPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(3) };
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
            
            lstShapes.Columns.Add("Z", 35);
            lstShapes.Columns.Add("Shape Name", 160);
            lstShapes.Columns.Add("Type", 80);
            lstShapes.Columns.Add("W", 45);
            lstShapes.Columns.Add("H", 45);
            
            lstShapes.SelectedIndexChanged += LstShapes_SelectedIndexChanged;
            lstShapes.AfterLabelEdit += LstShapes_AfterLabelEdit;
            lstShapes.ItemDrag += LstShapes_ItemDrag;
            lstShapes.DragEnter += LstShapes_DragEnter;
            lstShapes.DragDrop += LstShapes_DragDrop;
            lstShapes.DragOver += LstShapes_DragOver;
            lstShapes.DragLeave += LstShapes_DragLeave;
            
            mainLayout.Controls.Add(lstShapes, 0, 1);

            // Group: Layout & Distribution
            GroupBox grpAlign = new GroupBox { Text = "Layout & Distribution", Dock = DockStyle.Fill, Margin = new Padding(3) };
            FlowLayoutPanel flowAlign = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            grpAlign.Controls.Add(flowAlign);
            mainLayout.Controls.Add(grpAlign, 0, 2);

            btnAlignLeft = CreateToolButton("Left", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignLefts));
            btnAlignCenter = CreateToolButton("Center", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignCenters));
            btnAlignRight = CreateToolButton("Right", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignRights));
            btnAlignTop = CreateToolButton("Top", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignTops));
            btnAlignMiddle = CreateToolButton("Middle", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignMiddles));
            btnAlignBottom = CreateToolButton("Bottom", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignBottoms));
            btnDistributeH = CreateToolButton("Dist H", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeHorizontally));
            btnDistributeV = CreateToolButton("Dist V", (s, e) => DistributeSelected(Office.MsoDistributeCmd.msoDistributeVertically));
            btnCenterH = CreateToolButton("Ctr Slid H", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignCenters, true));
            btnCenterV = CreateToolButton("Ctr Slid V", (s, e) => AlignSelected(Office.MsoAlignCmd.msoAlignMiddles, true));

            flowAlign.Controls.AddRange(new Control[] { btnAlignLeft, btnAlignCenter, btnAlignRight, btnAlignTop, btnAlignMiddle, btnAlignBottom, btnDistributeH, btnDistributeV, btnCenterH, btnCenterV });

            // Group: Quick Draw Tools
            GroupBox grpDrawing = new GroupBox { Text = "Quick Draw Tools", Dock = DockStyle.Fill, Margin = new Padding(3) };
            FlowLayoutPanel flowDrawing = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            grpDrawing.Controls.Add(flowDrawing);
            mainLayout.Controls.Add(grpDrawing, 0, 3);

            btnLine = CreateToolButton("Line", (s, e) => ActivateDrawingTool(
                new[] { "ShapeStraightConnector" }, 32, "Line"));
            btnLine.BackColor = Color.AliceBlue;
            btnRect = CreateToolButton("Rect", (s, e) => ActivateDrawingTool(
                new[] { "ShapeRectangle" }, 0, "Rectangle"));
            btnRect.BackColor = Color.AliceBlue;
            btnFreeform = CreateToolButton("Freeform", (s, e) => ActivateDrawingTool(
                new[] { "ShapeFreeform", "ShapeFreeformShape", "Freeform", "FreeformTool" }, 200, "Freeform"));
            btnFreeform.BackColor = Color.AliceBlue;

            flowDrawing.Controls.AddRange(new Control[] { btnLine, btnRect, btnFreeform });

            // Group: Industrial Efficiency Tools
            GroupBox grpAdvanced = new GroupBox { Text = "Industrial Efficiency Tools", Dock = DockStyle.Fill, Margin = new Padding(3) };
            FlowLayoutPanel flowAdvanced = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, WrapContents = true };
            grpAdvanced.Controls.Add(flowAdvanced);
            mainLayout.Controls.Add(grpAdvanced, 0, 4);

            btnMatchWidth = CreateToolButton("Match W", (s, e) => MatchSize(true, false));
            btnMatchHeight = CreateToolButton("Match H", (s, e) => MatchSize(false, true));
            btnMatchBoth = CreateToolButton("Match All", (s, e) => MatchSize(true, true));
            btnMatchBoth.BackColor = Color.Ivory;
            
            btnStackH = CreateToolButton("Stack H", (s, e) => StackShapes(true));
            btnStackV = CreateToolButton("Stack V", (s, e) => StackShapes(false));
            btnRotate90 = CreateToolButton("Rot 90°", (s, e) => RotateSelected(90));

            btnSwap = CreateToolButton("Swap Pos", (s, e) => SwapShapes());
            btnSelectSameType = CreateToolButton("Same Type", (s, e) => SelectSameType());
            btnSelectAll = CreateToolButton("Sel All", (s, e) => SelectAllShapes());
            btnToFront = CreateToolButton("To Front", (s, e) => ZOrderExtreme(Office.MsoZOrderCmd.msoBringToFront));
            btnToBack = CreateToolButton("To Back", (s, e) => ZOrderExtreme(Office.MsoZOrderCmd.msoSendToBack));
            btnForward = CreateToolButton("Fwd Step", (s, e) => ZOrderExtreme(Office.MsoZOrderCmd.msoBringForward));
            btnBackward = CreateToolButton("Back Step", (s, e) => ZOrderExtreme(Office.MsoZOrderCmd.msoSendBackward));
            btnGroup = CreateToolButton("Group", (s, e) => GroupSelected());
            btnUngroup = CreateToolButton("Ungroup", (s, e) => UngroupSelected());
            btnHideAll = CreateToolButton("Hide All", (s, e) => ToggleAllVisibility(false));
            btnShowAll = CreateToolButton("Show All", (s, e) => ToggleAllVisibility(true));
            btnDelete = CreateToolButton("Delete", (s, e) => DeleteSelected());
            btnDelete.BackColor = Color.MistyRose;

            flowAdvanced.Controls.AddRange(new Control[] { btnMatchWidth, btnMatchHeight, btnMatchBoth, btnStackH, btnStackV, btnRotate90, btnSwap, btnSelectSameType, btnSelectAll, btnToFront, btnToBack, btnForward, btnBackward, btnGroup, btnUngroup, btnHideAll, btnShowAll, btnDelete });

            // Options Panel
            FlowLayoutPanel flowOptions = new FlowLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(3), FlowDirection = FlowDirection.LeftToRight };
            mainLayout.Controls.Add(flowOptions, 0, 5);

            chkFocusMode = new CheckBox { Text = "Focus", AutoSize = true, Margin = new Padding(5, 2, 5, 2) };
            chkFocusMode.CheckedChanged += (s, e) => ApplyVisibility();
            
            chkInverse = new CheckBox { Text = "Inverse", AutoSize = true, Margin = new Padding(5, 2, 5, 2) };
            chkInverse.CheckedChanged += (s, e) => ApplyVisibility();

            btnRefresh = new Button { Text = "Refresh", AutoSize = true, FlatStyle = FlatStyle.System, Margin = new Padding(5, 0, 5, 0) };
            btnRefresh.Click += (s, e) => LoadShapes();

            flowOptions.Controls.AddRange(new Control[] { chkFocusMode, chkInverse, btnRefresh });
        }

        private Button CreateToolButton(string text, EventHandler onClick)
        {
            Button btn = new Button { Text = text, Width = 60, Height = 28, Margin = new Padding(1), FlatStyle = FlatStyle.Flat, Font = new Font("Segoe UI", 7.5f) };
            btn.Click += onClick;
            return btn;
        }

        /// <summary>
        /// Prepares PowerPoint for drawing tool activation by validating prerequisites,
        /// switching to Normal view, activating the slide editing pane, and transferring
        /// focus from the WinForms window to the PowerPoint slide surface.
        /// Returns the active window if successful; null otherwise.
        /// </summary>
        private PowerPoint.DocumentWindow PrepareForDrawingTool(string toolName, out IntPtr pptHwnd, out bool attached)
        {
            pptHwnd = IntPtr.Zero;
            attached = false;

            // 1. Validate a presentation is open
            try
            {
                if (pptApp.Presentations.Count == 0)
                {
                    ShowDrawingToolError(toolName, "No presentation is open. Please open or create a presentation first.");
                    return null;
                }
            }
            catch
            {
                ShowDrawingToolError(toolName, "Cannot access PowerPoint. Please ensure it is running.");
                return null;
            }

            // 2. Validate active window
            PowerPoint.DocumentWindow activeWindow;
            try
            {
                activeWindow = pptApp.ActiveWindow;
                if (activeWindow == null) throw new InvalidOperationException();
            }
            catch
            {
                ShowDrawingToolError(toolName, "No active PowerPoint window found. Please open a presentation.");
                return null;
            }

            // 3. Ensure Normal or Slide view (required for drawing tools)
            try
            {
                var viewType = activeWindow.ViewType;
                if (viewType != PowerPoint.PpViewType.ppViewNormal &&
                    viewType != PowerPoint.PpViewType.ppViewSlide)
                {
                    activeWindow.ViewType = PowerPoint.PpViewType.ppViewNormal;
                    System.Threading.Thread.Sleep(100);
                }
            }
            catch
            {
                ShowDrawingToolError(toolName, "Cannot switch to Normal view. Please switch manually and try again.");
                return null;
            }

            // 4. Validate an active slide exists
            try
            {
                object slideObj = activeWindow.View.Slide;
                if (slideObj == null) throw new InvalidOperationException();
            }
            catch
            {
                ShowDrawingToolError(toolName, "No active slide found. Please select or add a slide first.");
                return null;
            }

            // 5. Get PowerPoint window handle
            try
            {
                pptHwnd = (IntPtr)activeWindow.HWND;
            }
            catch
            {
                ShowDrawingToolError(toolName, "Cannot access PowerPoint window handle.");
                return null;
            }

            // 6. Transfer focus to PowerPoint's slide editing pane
            uint currentThread = GetCurrentThreadId();
            uint pptThread = GetWindowThreadProcessId(pptHwnd, IntPtr.Zero);

            if (currentThread != pptThread)
            {
                attached = AttachThreadInput(currentThread, pptThread, true);
            }

            // Activate the slide editing pane (Pane 2 in Normal view)
            // This is the critical step — without it, drawing tools fail with E_FAIL
            try { activeWindow.Panes[2].Activate(); } catch { }

            // Bring PowerPoint window to foreground
            SetForegroundWindow(pptHwnd);
            BringWindowToTop(pptHwnd);
            SetFocus(pptHwnd);

            // Flush pending Windows messages and allow focus to fully stabilize
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(200);

            return activeWindow;
        }

        /// <summary>
        /// Re-activates the slide pane and sets foreground focus on PowerPoint.
        /// Used between retry attempts to re-establish the drawing context.
        /// </summary>
        private void RefocusPowerPoint(PowerPoint.DocumentWindow activeWindow, IntPtr pptHwnd, int delayMs)
        {
            try { activeWindow.Panes[2].Activate(); } catch { }
            SetForegroundWindow(pptHwnd);
            BringWindowToTop(pptHwnd);
            SetFocus(pptHwnd);
            System.Windows.Forms.Application.DoEvents();
            System.Threading.Thread.Sleep(delayMs);
        }

        /// <summary>
        /// Activates a PowerPoint drawing tool using a multi-strategy approach:
        ///   1. Try each idMso candidate via CommandBars.ExecuteMso (modern Ribbon API)
        ///   2. Fall back to CommandBars.FindControl with legacy control ID
        ///   3. Retry with extended delay on E_FAIL (focus/timing issue)
        /// This handles version differences across Office 2016/2019/2021/365.
        /// </summary>
        /// <param name="idMsoCandidates">Possible Ribbon idMso strings to try, in priority order.</param>
        /// <param name="legacyControlId">Legacy Office CommandBar control ID as fallback (0 to skip).</param>
        /// <param name="toolName">Human-readable tool name for error messages.</param>
        private void ActivateDrawingTool(string[] idMsoCandidates, int legacyControlId, string toolName)
        {
            IntPtr pptHwnd;
            bool attached;

            var activeWindow = PrepareForDrawingTool(toolName, out pptHwnd, out attached);
            if (activeWindow == null) return;

            try
            {
                // ---- Strategy 1: Try each idMso candidate (modern Ribbon API) ----
                foreach (var idMso in idMsoCandidates)
                {
                    try
                    {
                        pptApp.CommandBars.ExecuteMso(idMso);
                        return; // Success
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        // E_FAIL (0x80004005) = focus/timing issue → retry with delay
                        if (comEx.ErrorCode == unchecked((int)0x80004005))
                        {
                            try
                            {
                                System.Threading.Thread.Sleep(300);
                                RefocusPowerPoint(activeWindow, pptHwnd, 200);
                                pptApp.CommandBars.ExecuteMso(idMso);
                                return; // Retry succeeded
                            }
                            catch { /* retry also failed, continue to next candidate */ }
                        }
                        // For any other COM error (invalid arg, etc.), try next candidate
                        continue;
                    }
                    catch
                    {
                        continue; // Non-COM error, try next candidate
                    }
                }

                // ---- Strategy 2: Legacy CommandBars.FindControl fallback ----
                if (legacyControlId > 0)
                {
                    try
                    {
                        // Re-stabilize focus before legacy call
                        RefocusPowerPoint(activeWindow, pptHwnd, 150);

                        Office.CommandBarControl ctrl = pptApp.CommandBars.FindControl(Id: legacyControlId);
                        if (ctrl != null)
                        {
                            ctrl.Execute();
                            return; // Success via legacy
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException comEx)
                    {
                        if (comEx.ErrorCode == unchecked((int)0x80004005))
                        {
                            // E_FAIL on legacy → one more retry with longer delay
                            try
                            {
                                System.Threading.Thread.Sleep(400);
                                RefocusPowerPoint(activeWindow, pptHwnd, 300);

                                Office.CommandBarControl ctrlRetry = pptApp.CommandBars.FindControl(Id: legacyControlId);
                                if (ctrlRetry != null)
                                {
                                    ctrlRetry.Execute();
                                    return; // Legacy retry succeeded
                                }
                            }
                            catch { /* final retry also failed */ }
                        }
                    }
                    catch { /* non-COM legacy error */ }
                }

                // ---- All strategies exhausted ----
                ShowDrawingToolError(toolName,
                    "The tool could not be activated in this PowerPoint version.\n\n" +
                    "Please try:\n" +
                    "\u2022 Click on the slide area in PowerPoint first, then try again\n" +
                    "\u2022 Ensure you are in Normal view (View \u2192 Normal)\n" +
                    "\u2022 Use Insert \u2192 Shapes from PowerPoint's ribbon directly");
            }
            finally
            {
                if (attached)
                {
                    uint currentThread = GetCurrentThreadId();
                    uint pptThread = GetWindowThreadProcessId(pptHwnd, IntPtr.Zero);
                    AttachThreadInput(currentThread, pptThread, false);
                }
            }
        }

        /// <summary>
        /// Displays a standardized error message for drawing tool failures.
        /// </summary>
        private void ShowDrawingToolError(string toolName, string message)
        {
            MessageBox.Show(message, $"{toolName} \u2014 Drawing Tool",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private void ShapeManagerForm_Load(object sender, EventArgs e) { SetPlaceholder(txtSearch, "Search shapes by name..."); LoadShapes(); }
        private void ShapeManagerForm_FormClosing(object sender, FormClosingEventArgs e) { RestoreOriginalVisibility(); }

        public void SyncSelectionFromPowerPoint(PowerPoint.Selection sel)
        {
            if (isInternalChange) return;
            if (this.InvokeRequired) { this.Invoke(new Action(() => SyncSelectionFromPowerPoint(sel))); return; }
            try {
                isInternalChange = true;
                lstShapes.SelectedIndexChanged -= LstShapes_SelectedIndexChanged;
                HashSet<int> selectedIds = new HashSet<int>();
                if (sel != null && sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { foreach (PowerPoint.Shape s in sel.ShapeRange) try { selectedIds.Add(s.Id); } catch { } }
                lstShapes.BeginUpdate();
                foreach (ListViewItem item in lstShapes.Items) { if (item.Tag is PowerPoint.Shape s) { try { bool selected = selectedIds.Contains(s.Id); if (item.Selected != selected) item.Selected = selected; } catch { } } }
                if (lstShapes.SelectedItems.Count > 0) lstShapes.SelectedItems[0].EnsureVisible();
                lstShapes.EndUpdate();
            } catch { } finally { lstShapes.SelectedIndexChanged += LstShapes_SelectedIndexChanged; isInternalChange = false; }
        }

        private void FilterShapes() { string filter = txtSearch.Text.ToLower(); lstShapes.BeginUpdate(); foreach (ListViewItem item in lstShapes.Items) { if (string.IsNullOrEmpty(filter)) item.BackColor = SystemColors.Window; else if (item.SubItems[1].Text.ToLower().Contains(filter)) item.BackColor = Color.LightYellow; else item.BackColor = SystemColors.Window; } lstShapes.EndUpdate(); }
        private void LoadShapes() { lstShapes.BeginUpdate(); lstShapes.Items.Clear(); try { var slide = GetActiveSlide(); if (slide == null) return; for (int i = slide.Shapes.Count; i >= 1; i--) { try { PowerPoint.Shape s = slide.Shapes[i]; ListViewItem item = new ListViewItem(i.ToString()); item.SubItems.Add(s.Name); item.SubItems.Add(s.Type.ToString().Replace("mso", "")); item.SubItems.Add(Math.Round(s.Width, 1).ToString()); item.SubItems.Add(Math.Round(s.Height, 1).ToString()); item.Tag = s; lstShapes.Items.Add(item); } catch { continue; } } } catch { } finally { lstShapes.EndUpdate(); } }
        private void AlignSelected(Office.MsoAlignCmd alignCmd, bool relativeToSlideForce = false) { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { var range = pptApp.ActiveWindow.Selection.ShapeRange; if (range.Count > 0) { Office.MsoTriState rel = (range.Count == 1 || relativeToSlideForce) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; range.Align(alignCmd, rel); } } } catch { } }
        private void DistributeSelected(Office.MsoDistributeCmd distCmd) { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { var range = pptApp.ActiveWindow.Selection.ShapeRange; if (range.Count >= 2) range.Distribute(distCmd, Office.MsoTriState.msoFalse); } } catch { } }
        private void ZOrderExtreme(Office.MsoZOrderCmd cmd) { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { pptApp.ActiveWindow.Selection.ShapeRange.ZOrder(cmd); LoadShapes(); } } catch { } }
        private void MatchSize(bool width, bool height) { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { var range = pptApp.ActiveWindow.Selection.ShapeRange; if (range.Count < 2) return; float refW = range[1].Width, refH = range[1].Height; for (int i = 2; i <= range.Count; i++) { if (width) range[i].Width = refW; if (height) range[i].Height = refH; } LoadShapes(); } } catch { } }
        private void SwapShapes() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { var range = pptApp.ActiveWindow.Selection.ShapeRange; if (range.Count != 2) return; float tL = range[1].Left, tT = range[1].Top; range[1].Left = range[2].Left; range[1].Top = range[2].Top; range[2].Left = tL; range[2].Top = tT; } } catch { } }
        private void SelectSameType() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { var type = pptApp.ActiveWindow.Selection.ShapeRange[1].Type; var slide = GetActiveSlide(); if (slide == null) return; List<string> names = new List<string>(); foreach (PowerPoint.Shape s in slide.Shapes) if (s.Type == type) names.Add(s.Name); if (names.Count > 0) { isInternalChange = true; slide.Shapes.Range(names.ToArray()).Select(); isInternalChange = false; SyncSelectionFromPowerPoint(pptApp.ActiveWindow.Selection); } } } catch { } }
        private void SelectAllShapes() { try { var slide = GetActiveSlide(); if (slide == null || slide.Shapes.Count == 0) return; isInternalChange = true; slide.Shapes.SelectAll(); isInternalChange = false; SyncSelectionFromPowerPoint(pptApp.ActiveWindow.Selection); } catch { } }
        private void DeleteSelected() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { if (MessageBox.Show("Delete selected items?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes) { isInternalChange = true; pptApp.ActiveWindow.Selection.Delete(); isInternalChange = false; LoadShapes(); } } } catch { } }
        private void ToggleAllVisibility(bool visible) { try { var slide = GetActiveSlide(); if (slide == null) return; foreach (PowerPoint.Shape s in slide.Shapes) try { s.Visible = visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { } LoadShapes(); } catch { } }
        private void GroupSelected() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { pptApp.ActiveWindow.Selection.ShapeRange.Group(); LoadShapes(); } } catch { } }
        private void UngroupSelected() { try { if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes) { pptApp.ActiveWindow.Selection.ShapeRange.Ungroup(); LoadShapes(); } } catch { } }
        private void StackShapes(bool horizontal)
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var range = pptApp.ActiveWindow.Selection.ShapeRange;
                    if (range.Count < 2) return;

                    List<PowerPoint.Shape> sorted = new List<PowerPoint.Shape>();
                    foreach (PowerPoint.Shape s in range) sorted.Add(s);
                    
                    if (horizontal)
                        sorted.Sort((a, b) => a.Left.CompareTo(b.Left));
                    else
                        sorted.Sort((a, b) => a.Top.CompareTo(b.Top));

                    float currentPos = horizontal ? sorted[0].Left + sorted[0].Width : sorted[0].Top + sorted[0].Height;

                    for (int i = 1; i < sorted.Count; i++)
                    {
                        if (horizontal)
                        {
                            sorted[i].Left = currentPos;
                            currentPos += sorted[i].Width;
                        }
                        else
                        {
                            sorted[i].Top = currentPos;
                            currentPos += sorted[i].Height;
                        }
                    }
                    LoadShapes();
                }
            }
            catch { }
        }

        private void RotateSelected(float degrees)
        {
            try
            {
                if (pptApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    pptApp.ActiveWindow.Selection.ShapeRange.IncrementRotation(degrees);
                    LoadShapes();
                }
            }
            catch { }
        }

        private void LstShapes_AfterLabelEdit(object sender, LabelEditEventArgs e) { if (e.Label == null) return; try { ListViewItem item = lstShapes.Items[e.Item]; if (item.Tag is PowerPoint.Shape s) s.Name = e.Label; } catch { e.CancelEdit = true; } }

        private PowerPoint.Slide GetActiveSlide()
        {
            try {
                if (pptApp.ActiveWindow == null) return null;
                try { return (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide; } catch { }
                try { if (pptApp.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionNone) return (PowerPoint.Slide)pptApp.ActiveWindow.Selection.SlideRange[1]; } catch { }
            } catch { }
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
            try {
                if (lstShapes.SelectedItems.Count == 0) { pptApp.ActiveWindow.Selection.Unselect(); return; }
                List<string> names = new List<string>();
                foreach (ListViewItem item in lstShapes.SelectedItems) if (item.Tag is PowerPoint.Shape s) names.Add(s.Name);
                var slide = GetActiveSlide();
                if (slide != null && names.Count > 0) slide.Shapes.Range(names.ToArray()).Select();
            } catch { }
        }

        private void ApplyVisibility()
        {
            if (!chkFocusMode.Checked) { RestoreOriginalVisibility(); return; }
            var slide = GetActiveSlide();
            if (slide == null) return;
            HashSet<int> selectedIds = new HashSet<int>();
            foreach (ListViewItem item in lstShapes.SelectedItems) if (item.Tag is PowerPoint.Shape s) selectedIds.Add(s.Id);
            foreach (PowerPoint.Shape s in slide.Shapes) {
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
            if (targetItem != null) {
                int targetIndex = targetItem.Index;
                Rectangle itemBounds = targetItem.GetBounds(ItemBoundsPortion.Entire);
                lstShapes.InsertionMark.AppearsAfterItem = (cp.Y > itemBounds.Top + (itemBounds.Height / 2));
                lstShapes.InsertionMark.Index = targetIndex;
                e.Effect = DragDropEffects.Move;
            } else { e.Effect = DragDropEffects.None; lstShapes.InsertionMark.Index = -1; }
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
