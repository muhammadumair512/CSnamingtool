using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace name_tool
{
    /// <summary>
    /// Provides Photoshop-powered Quick Selection for PowerPoint images.
    /// Exports a selected image to Adobe Photoshop, activates the Quick Selection tool,
    /// lets the user interactively select an area, then imports the cutout back into
    /// PowerPoint as an independent shape that can be animated or manipulated.
    /// </summary>
    public static class QuickSelectHelper
    {
        /// <summary>
        /// Main entry point — called from the Ribbon button.
        /// </summary>
        public static void Execute(PowerPoint.Application pptApp)
        {
            // ── 1. Validate PowerPoint state ──────────────────────────────────
            PowerPoint.Selection sel;
            try { sel = pptApp.ActiveWindow.Selection; }
            catch
            {
                ShowError("No active PowerPoint window found.\nPlease open a presentation first.");
                return;
            }

            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                ShowError("Please select an image or shape on the slide first,\nthen click Quick Select.");
                return;
            }

            PowerPoint.Shape shape = sel.ShapeRange[1];

            // Store original shape properties for position mapping
            float pptLeft = shape.Left;
            float pptTop = shape.Top;
            float pptWidth = shape.Width;
            float pptHeight = shape.Height;
            float pptRotation = shape.Rotation;

            // ── 2. Create temp directory and export image ─────────────────────
            string tempDir = Path.Combine(Path.GetTempPath(),
                "PPTQuickSel_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(tempDir);

            string exportPath = Path.Combine(tempDir, "source.png");
            string cutoutPath = Path.Combine(tempDir, "cutout.png");

            try
            {
                // Export the selected shape as a PNG image
                string exportError = "";
                bool exported = false;

                // Strategy A: Direct shape export via SaveAs temp file
                if (!exported)
                    exported = ExportShapeDirectly(pptApp, shape, exportPath, out exportError);

                // Strategy B: Temp presentation with visible window
                if (!exported)
                    exported = ExportShapeViaSlide(pptApp, shape, exportPath, out exportError);

                // Strategy C: Export the full active slide, then crop to shape bounds
                if (!exported)
                    exported = ExportShapeViaSlideCrop(pptApp, shape, exportPath, out exportError);

                // Strategy D: Clipboard (STA-safe)
                if (!exported)
                    exported = ExportShapeViaClipboard(shape, exportPath, out exportError);

                if (!exported)
                {
                    ShowError("Failed to export the selected shape as an image.\n\n" +
                              "Details: " + exportError + "\n\n" +
                              "Try right-clicking the image in PowerPoint → 'Save as Picture',\n" +
                              "then use File → Open in Photoshop directly.");
                    return;
                }

                // ── 3. Two-phase Photoshop workflow (no COM, non-blocking) ──
                //   Phase 1: Launch PS with a setup script that opens the image
                //            and activates Quick Selection Tool, then exits.
                //   User freely works in Photoshop (no modal dialog blocking).
                //   Phase 2: When user clicks "Import" in our C# dialog, we
                //            launch a second JSX script that reads the selection,
                //            processes the cutout, saves PNG, and writes result.

                string psExePath = FindPhotoshopExe();
                if (psExePath == null)
                {
                    ShowError("Could not find Adobe Photoshop.\n\n" +
                              "Please ensure Photoshop is installed at:\n" +
                              "  C:\\Program Files\\Adobe\\Adobe Photoshop 2026\\\n\n" +
                              "Or any standard Adobe installation path.");
                    return;
                }

                string setupJsxPath   = Path.Combine(tempDir, "setup.jsx");
                string processJsxPath = Path.Combine(tempDir, "process.jsx");
                string resultFilePath = Path.Combine(tempDir, "result.txt");

                // ── Phase 1: Open image + activate Quick Selection ────────────
                string setupScript = BuildSetupJsx(exportPath.Replace("\\", "/"));
                File.WriteAllText(setupJsxPath, setupScript, System.Text.Encoding.UTF8);

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = psExePath,
                        Arguments = "\"" + setupJsxPath + "\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowError("Failed to launch Photoshop.\n\n" + ex.Message);
                    return;
                }

                // ── Show C# control panel (user works freely in PS) ───────────
                DialogResult dlgResult;
                using (Form controlPanel = BuildControlPanel())
                {
                    dlgResult = controlPanel.ShowDialog();
                }

                if (dlgResult != DialogResult.OK)
                {
                    // User cancelled — try to close PS document via a cleanup script
                    string cleanupScript = BuildCleanupJsx();
                    string cleanupPath = Path.Combine(tempDir, "cleanup.jsx");
                    File.WriteAllText(cleanupPath, cleanupScript, System.Text.Encoding.UTF8);
                    try
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = psExePath,
                            Arguments = "\"" + cleanupPath + "\"",
                            UseShellExecute = true
                        });
                    }
                    catch { }
                    return;
                }

                // ── Phase 2: Process selection and export cutout ──────────────
                string processScript = BuildProcessJsx(
                    cutoutPath.Replace("\\", "/"),
                    resultFilePath.Replace("\\", "/"));
                File.WriteAllText(processJsxPath, processScript, System.Text.Encoding.UTF8);

                try
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = psExePath,
                        Arguments = "\"" + processJsxPath + "\"",
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    ShowError("Failed to run Photoshop export script.\n\n" + ex.Message);
                    return;
                }

                // ── Wait for result file from Phase 2 ─────────────────────────
                string scriptResult = WaitForResultFile(resultFilePath, 60);

                if (scriptResult == null)
                {
                    ShowError("Timed out waiting for Photoshop to process the selection.\n\n" +
                              "Please try again. Make sure Photoshop is not showing any dialogs.");
                    return;
                }

                // ── Handle script result ──────────────────────────────────────
                if (scriptResult == "NO_SELECTION")
                {
                    ShowError("No selection was detected in Photoshop.\n\n" +
                              "Please make a selection using the Quick Selection Tool\n" +
                              "before clicking \"Import Selection\".");
                    return;
                }

                if (scriptResult.StartsWith("ERROR:", StringComparison.OrdinalIgnoreCase))
                {
                    ShowError("Photoshop returned an error:\n" + scriptResult.Substring(6));
                    return;
                }

                // ── Parse bounds and compute PPT position ─────────────────────
                string[] parts = scriptResult.Split(',');
                if (parts.Length < 6)
                {
                    ShowError("Unexpected response from Photoshop:\n" + scriptResult);
                    return;
                }

                float selLeft  = float.Parse(parts[0], CultureInfo.InvariantCulture);
                float selTop   = float.Parse(parts[1], CultureInfo.InvariantCulture);
                float selRight = float.Parse(parts[2], CultureInfo.InvariantCulture);
                float selBot   = float.Parse(parts[3], CultureInfo.InvariantCulture);
                float imgW     = float.Parse(parts[4], CultureInfo.InvariantCulture);
                float imgH     = float.Parse(parts[5], CultureInfo.InvariantCulture);

                float scaleX = pptWidth / imgW;
                float scaleY = pptHeight / imgH;

                float newLeft   = pptLeft + selLeft * scaleX;
                float newTop    = pptTop  + selTop  * scaleY;
                float newWidth  = (selRight - selLeft) * scaleX;
                float newHeight = (selBot   - selTop)  * scaleY;

                // ── Import cutout into PowerPoint ─────────────────────────────
                if (!File.Exists(cutoutPath))
                {
                    ShowError("The cutout image was not created by Photoshop.\n" +
                              "Please try the operation again.");
                    return;
                }

                PowerPoint.Slide slide;
                try { slide = (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide; }
                catch
                {
                    ShowError("Cannot access the active slide.\nPlease make sure a slide is selected.");
                    return;
                }

                PowerPoint.Shape newShape = slide.Shapes.AddPicture(
                    cutoutPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    newLeft, newTop, newWidth, newHeight);

                newShape.Name = "QuickSelect_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (Math.Abs(pptRotation) > 0.01f)
                    newShape.Rotation = pptRotation;

                try
                {
                    int origZ = shape.ZOrderPosition;
                    while (newShape.ZOrderPosition > origZ + 1)
                        newShape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
                }
                catch { }

                try { newShape.Select(); } catch { }

                MessageBox.Show(
                    "Quick Select completed successfully!\n\n" +
                    "A new shape \"" + newShape.Name + "\" has been created\n" +
                    "from your Photoshop selection.\n\n" +
                    "You can now:\n" +
                    "  \u2022  Add animations (Animations tab)\n" +
                    "  \u2022  Move, resize, or rotate it\n" +
                    "  \u2022  Apply effects and transitions\n" +
                    "  \u2022  Layer it with other shapes",
                    "Quick Select \u2014 Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            finally
            {
                // Clean up temp files (best-effort, delayed)
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    Thread.Sleep(5000);
                    try { Directory.Delete(tempDir, true); } catch { }
                });
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  Export helpers
        // ═══════════════════════════════════════════════════════════════════

        // ═══════════════════════════════════════════════════════════════════
        //  Export strategies (4 fallbacks)
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Strategy A: Export shape picture data directly.
        /// Uses shape.Export (available on Picture/OLE shapes).
        /// Falls back to saving the shape as a temp picture via PPT interop.
        /// </summary>
        private static bool ExportShapeDirectly(PowerPoint.Application pptApp,
            PowerPoint.Shape shape, string filePath, out string error)
        {
            error = "";
            try
            {
                // Use the shape's parent slide to export just this shape
                // by grouping single-select copy/paste approach
                shape.Copy();
                Thread.Sleep(200);

                // Create a temp presentation WITH a window (critical for paste to work)
                PowerPoint.Presentation tempPres = null;
                try
                {
                    tempPres = pptApp.Presentations.Add(Office.MsoTriState.msoTrue);

                    // Set slide to shape dimensions
                    tempPres.PageSetup.SlideWidth = shape.Width;
                    tempPres.PageSetup.SlideHeight = shape.Height;

                    var tempSlide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    tempSlide.FollowMasterBackground = Office.MsoTriState.msoFalse;

                    // Set white background (ensures clean export)
                    tempSlide.Background.Fill.Visible = Office.MsoTriState.msoTrue;
                    tempSlide.Background.Fill.Solid();
                    tempSlide.Background.Fill.ForeColor.RGB = 0xFFFFFF;

                    // Paste using PasteSpecial for maximum reliability
                    PowerPoint.ShapeRange pasted = null;
                    try
                    {
                        pasted = tempSlide.Shapes.PasteSpecial(
                            PowerPoint.PpPasteDataType.ppPastePNG);
                    }
                    catch
                    {
                        try
                        {
                            Thread.Sleep(200);
                            pasted = tempSlide.Shapes.Paste();
                        }
                        catch (Exception pasteEx)
                        {
                            error = "Paste failed: " + pasteEx.Message;
                            return false;
                        }
                    }

                    if (pasted == null || pasted.Count == 0)
                    {
                        error = "Nothing was pasted into temp slide.";
                        return false;
                    }

                    PowerPoint.Shape pastedShape = pasted[1];
                    pastedShape.Left = 0;
                    pastedShape.Top = 0;
                    pastedShape.Width = shape.Width;
                    pastedShape.Height = shape.Height;
                    pastedShape.Rotation = 0;

                    int pixW = Math.Max(1, (int)(shape.Width * 96.0f / 72.0f * 2));
                    int pixH = Math.Max(1, (int)(shape.Height * 96.0f / 72.0f * 2));

                    tempSlide.Export(filePath, "PNG", pixW, pixH);

                    return File.Exists(filePath);
                }
                finally
                {
                    if (tempPres != null)
                    {
                        try { tempPres.Close(); } catch { }
                        try { Marshal.ReleaseComObject(tempPres); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                error = "Direct export: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Strategy B: Creates a hidden temp presentation with msoFalse window.
        /// Some PPT versions allow this; kept as fallback.
        /// </summary>
        private static bool ExportShapeViaSlide(PowerPoint.Application pptApp,
                                                 PowerPoint.Shape shape, string filePath,
                                                 out string error)
        {
            error = "";
            PowerPoint.Presentation tempPres = null;
            try
            {
                tempPres = pptApp.Presentations.Add(Office.MsoTriState.msoFalse);

                tempPres.PageSetup.SlideWidth = shape.Width;
                tempPres.PageSetup.SlideHeight = shape.Height;

                var tempSlide = tempPres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                tempSlide.FollowMasterBackground = Office.MsoTriState.msoFalse;
                tempSlide.Background.Fill.Visible = Office.MsoTriState.msoFalse;

                shape.Copy();
                Thread.Sleep(400);

                PowerPoint.ShapeRange pasted = tempSlide.Shapes.Paste();
                PowerPoint.Shape pastedShape = pasted[1];

                pastedShape.Left = 0;
                pastedShape.Top = 0;
                pastedShape.Width = shape.Width;
                pastedShape.Height = shape.Height;
                pastedShape.Rotation = 0;

                int pixW = Math.Max(1, (int)(shape.Width * 96.0f / 72.0f * 2));
                int pixH = Math.Max(1, (int)(shape.Height * 96.0f / 72.0f * 2));

                tempSlide.Export(filePath, "PNG", pixW, pixH);

                return File.Exists(filePath);
            }
            catch (Exception ex)
            {
                error = "SlideExport(hidden): " + ex.Message;
                return false;
            }
            finally
            {
                if (tempPres != null)
                {
                    try { tempPres.Close(); } catch { }
                    try { Marshal.ReleaseComObject(tempPres); } catch { }
                }
            }
        }

        /// <summary>
        /// Strategy C: Export the FULL active slide as PNG at high res, then crop
        /// to the shape's bounding box in GDI+. Works even when copy/paste fails.
        /// </summary>
        private static bool ExportShapeViaSlideCrop(PowerPoint.Application pptApp,
            PowerPoint.Shape shape, string filePath, out string error)
        {
            error = "";
            string fullSlidePath = Path.Combine(Path.GetDirectoryName(filePath), "fullslide_temp.png");
            try
            {
                PowerPoint.Slide slide = (PowerPoint.Slide)pptApp.ActiveWindow.View.Slide;

                float slideW = pptApp.ActivePresentation.PageSetup.SlideWidth;
                float slideH = pptApp.ActivePresentation.PageSetup.SlideHeight;

                // Export full slide at 2× resolution
                int exportW = Math.Max(1, (int)(slideW * 96.0f / 72.0f * 2));
                int exportH = Math.Max(1, (int)(slideH * 96.0f / 72.0f * 2));

                slide.Export(fullSlidePath, "PNG", exportW, exportH);

                if (!File.Exists(fullSlidePath))
                {
                    error = "Full slide export produced no file.";
                    return false;
                }

                // Calculate the shape's bounding box in pixel coordinates
                float scaleX = exportW / slideW;
                float scaleY = exportH / slideH;

                int cropX = Math.Max(0, (int)(shape.Left * scaleX));
                int cropY = Math.Max(0, (int)(shape.Top * scaleY));
                int cropW = Math.Max(1, (int)(shape.Width * scaleX));
                int cropH = Math.Max(1, (int)(shape.Height * scaleY));

                using (Image fullImg = Image.FromFile(fullSlidePath))
                {
                    // Clamp to image bounds
                    if (cropX + cropW > fullImg.Width) cropW = fullImg.Width - cropX;
                    if (cropY + cropH > fullImg.Height) cropH = fullImg.Height - cropY;
                    if (cropW <= 0 || cropH <= 0)
                    {
                        error = "Shape bounds outside slide area.";
                        return false;
                    }

                    using (Bitmap bmp = new Bitmap(cropW, cropH))
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(fullImg,
                            new Rectangle(0, 0, cropW, cropH),
                            new Rectangle(cropX, cropY, cropW, cropH),
                            GraphicsUnit.Pixel);
                        bmp.Save(filePath, ImageFormat.Png);
                    }
                }

                return File.Exists(filePath);
            }
            catch (Exception ex)
            {
                error = "SlideCrop: " + ex.Message;
                return false;
            }
            finally
            {
                try { if (File.Exists(fullSlidePath)) File.Delete(fullSlidePath); } catch { }
            }
        }

        /// <summary>
        /// Strategy D: Clipboard-based export (STA-safe).
        /// </summary>
        private static bool ExportShapeViaClipboard(PowerPoint.Shape shape, string filePath,
            out string error)
        {
            error = "";
            try
            {
                shape.Copy();

                // Clipboard must be accessed from an STA thread
                Image clipImage = null;
                Exception threadEx = null;

                Thread staThread = new Thread(() =>
                {
                    try
                    {
                        for (int attempt = 0; attempt < 15; attempt++)
                        {
                            Thread.Sleep(200);
                            try
                            {
                                if (Clipboard.ContainsImage())
                                {
                                    clipImage = Clipboard.GetImage();
                                    if (clipImage != null) break;
                                }
                            }
                            catch { /* clipboard busy */ }
                        }
                    }
                    catch (Exception ex) { threadEx = ex; }
                });
                staThread.SetApartmentState(ApartmentState.STA);
                staThread.Start();
                staThread.Join(8000); // 8 second timeout

                if (threadEx != null)
                {
                    error = "Clipboard thread: " + threadEx.Message;
                    return false;
                }

                if (clipImage == null)
                {
                    error = "No image found on clipboard after copy.";
                    return false;
                }

                using (clipImage)
                using (var bmp = new Bitmap(clipImage))
                {
                    bmp.Save(filePath, ImageFormat.Png);
                }

                return File.Exists(filePath);
            }
            catch (Exception ex)
            {
                error = "Clipboard: " + ex.Message;
                return false;
            }
        }

        // ═══════════════════════════════════════════════════════════════════
        //  UI helpers
        // ═══════════════════════════════════════════════════════════════════

        private static void ShowError(string message)
        {
            MessageBox.Show(message, "Quick Select",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // ═══════════════════════════════════════════════════════════════════
        //  JSX Scripts (two-phase, non-blocking)
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Phase 1 JSX: Opens the source image and activates Quick Selection Tool.
        /// Script completes immediately — no modal dialog, user can work freely.
        /// </summary>
        private static string BuildSetupJsx(string sourcePath)
        {
            return @"
#target photoshop
app.bringToFront();
try {
    var srcFile = new File('__SOURCE__');
    if (srcFile.exists) {
        var doc = app.open(srcFile);
        // Activate Quick Selection Tool
        try {
            var desc = new ActionDescriptor();
            var ref = new ActionReference();
            ref.putClass(stringIDToTypeID('quickSelectTool'));
            desc.putReference(charIDToTypeID('null'), ref);
            executeAction(charIDToTypeID('setd'), desc, DialogModes.NO);
        } catch(e) {
            try {
                var desc2 = new ActionDescriptor();
                var ref2 = new ActionReference();
                ref2.putClass(stringIDToTypeID('magicWandTool'));
                desc2.putReference(charIDToTypeID('null'), ref2);
                executeAction(charIDToTypeID('setd'), desc2, DialogModes.NO);
            } catch(e2) {}
        }
    }
} catch(e) {}
".Replace("__SOURCE__", sourcePath);
        }

        /// <summary>
        /// Phase 2 JSX: Reads the current selection, processes the cutout,
        /// saves transparent PNG, writes result coordinates to file.
        /// </summary>
        private static string BuildProcessJsx(string cutoutPath, string resultPath)
        {
            return @"
#target photoshop
function writeResult(text) {
    var f = new File('__RESULT__');
    var folder = f.parent;
    if (!folder.exists) folder.create();
    f.encoding = 'UTF-8';
    f.open('w');
    f.write(text);
    f.close();
}
try {
    if (app.documents.length === 0) {
        writeResult('ERROR:No document is open in Photoshop.');
    } else {
        var doc = app.activeDocument;
        var bounds = null;
        try { bounds = doc.selection.bounds; } catch(e) {}

        if (bounds === null) {
            writeResult('NO_SELECTION');
        } else {
            var selL = bounds[0].as('px');
            var selT = bounds[1].as('px');
            var selR = bounds[2].as('px');
            var selB = bounds[3].as('px');
            var docW = doc.width.as('px');
            var docH = doc.height.as('px');

            // Store the selection into a named alpha channel (survives duplicate + flatten)
            var selChannel = doc.channels.add();
            selChannel.name = 'PPTQuickSelMask';
            selChannel.kind = ChannelType.SELECTEDAREA;
            doc.selection.store(selChannel);

            // Duplicate the entire document (includes alpha channels)
            var dupDoc = doc.duplicate('QuickSelectCutout');

            // Ensure RGB mode (required for proper transparency in PNG)
            if (dupDoc.mode !== DocumentMode.RGB) {
                dupDoc.changeMode(ChangeMode.RGB);
            }

            // Flatten all layers into a single Background layer
            dupDoc.flatten();

            // Convert Background to regular layer (Background cannot have transparency)
            dupDoc.activeLayer.isBackgroundLayer = false;
            dupDoc.activeLayer.name = 'Cutout';

            // Find the stored selection mask in the duplicate
            var maskChan = null;
            for (var c = 0; c < dupDoc.channels.length; c++) {
                if (dupDoc.channels[c].name === 'PPTQuickSelMask') {
                    maskChan = dupDoc.channels[c];
                    break;
                }
            }

            if (maskChan === null) {
                writeResult('ERROR:Selection mask channel not found after duplicate.');
                dupDoc.close(SaveOptions.DONOTSAVECHANGES);
            } else {
                // Load the stored selection back as marching ants
                dupDoc.selection.load(maskChan, SelectionType.REPLACE);
                maskChan.remove();

                // ╔═══════════════════════════════════════════════════════════╗
                // ║  CREATE TRANSPARENCY VIA LAYER MASK                      ║
                // ║  This is the professional Photoshop method:              ║
                // ║  - Layer mask from selection = white(visible)/black(hide)║
                // ║  - Apply mask = bake transparency into actual pixels     ║
                // ╚═══════════════════════════════════════════════════════════╝
                var usedMask = false;
                try {
                    // Create layer mask: ""Reveal Selection""
                    // Selected area → white (visible), rest → black (hidden)
                    var mkDesc = new ActionDescriptor();
                    mkDesc.putClass(charIDToTypeID('Nw  '), charIDToTypeID('Chnl'));
                    var mkRef = new ActionReference();
                    mkRef.putEnumerated(
                        charIDToTypeID('Chnl'),
                        charIDToTypeID('Chnl'),
                        charIDToTypeID('Msk '));
                    mkDesc.putReference(charIDToTypeID('At  '), mkRef);
                    mkDesc.putEnumerated(
                        charIDToTypeID('Usng'),
                        charIDToTypeID('UsrM'),
                        charIDToTypeID('RvlS'));
                    executeAction(charIDToTypeID('Mk  '), mkDesc, DialogModes.NO);

                    // Apply (flatten) the layer mask into the layer pixels
                    // Hidden areas become truly transparent pixels
                    var apDesc = new ActionDescriptor();
                    var apRef = new ActionReference();
                    apRef.putEnumerated(
                        charIDToTypeID('Chnl'),
                        charIDToTypeID('Chnl'),
                        charIDToTypeID('Msk '));
                    apDesc.putReference(charIDToTypeID('null'), apRef);
                    apDesc.putBoolean(charIDToTypeID('Aply'), true);
                    executeAction(charIDToTypeID('Dlt '), apDesc, DialogModes.NO);

                    usedMask = true;
                } catch(maskErr) {
                    // Layer mask failed — fallback to select-invert-delete
                    usedMask = false;
                }

                if (!usedMask) {
                    // Fallback: reload selection, invert, delete
                    // Re-add the alpha channel from original doc
                    try {
                        // We already removed maskChan, so re-select
                        // The selection should still be active from the load above
                        dupDoc.selection.invert();
                        dupDoc.activeLayer.clear();
                    } catch(clearErr) {
                        writeResult('ERROR:Could not create transparency: ' + clearErr.message);
                        dupDoc.close(SaveOptions.DONOTSAVECHANGES);
                        doc.close(SaveOptions.DONOTSAVECHANGES);
                    }
                }

                // Deselect all
                try { dupDoc.selection.deselect(); } catch(ds) {}

                // Crop to the original selection bounding box
                dupDoc.crop([
                    new UnitValue(selL, 'px'),
                    new UnitValue(selT, 'px'),
                    new UnitValue(selR, 'px'),
                    new UnitValue(selB, 'px')
                ]);

                // ╔═══════════════════════════════════════════════════════════╗
                // ║  SAVE AS TRANSPARENT PNG-24                              ║
                // ║  Strategy 1: Save For Web (explicit transparency flag)   ║
                // ║  Strategy 2: saveAs with PNGSaveOptions (fallback)       ║
                // ╚═══════════════════════════════════════════════════════════╝
                var outFile = new File('__CUTOUT__');
                var outFolder = outFile.parent;
                if (!outFolder.exists) outFolder.create();

                var saved = false;

                // Strategy 1: Save For Web — guarantees PNG-24 alpha channel
                if (!saved) {
                    try {
                        var sfwOpts = new ExportOptionsSaveForWeb();
                        sfwOpts.format = SaveDocumentType.PNG;
                        sfwOpts.PNG8 = false;
                        sfwOpts.transparency = true;
                        sfwOpts.includeProfile = false;
                        sfwOpts.optimized = true;
                        dupDoc.exportDocument(outFile, ExportType.SAVEFORWEB, sfwOpts);
                        saved = outFile.exists;
                    } catch(sfwErr) {
                        // Save For Web failed — try next strategy
                    }
                }

                // Strategy 2: Standard saveAs with PNGSaveOptions
                if (!saved) {
                    try {
                        var pngOpts = new PNGSaveOptions();
                        pngOpts.compression = 6;
                        pngOpts.interlaced = false;
                        dupDoc.saveAs(outFile, pngOpts, true, Extension.LOWERCASE);
                        saved = true;
                    } catch(pngErr) {}
                }

                // Strategy 3: Save as TIFF with transparency then convert
                if (!saved) {
                    try {
                        var tiffFile = new File(outFile.fsName.replace(/\.png$/i, '.tif'));
                        var tiffOpts = new TiffSaveOptions();
                        tiffOpts.alphaChannels = true;
                        tiffOpts.transparency = true;
                        tiffOpts.layers = false;
                        dupDoc.saveAs(tiffFile, tiffOpts, true, Extension.LOWERCASE);
                        // Re-open and save as PNG
                        var tDoc = app.open(tiffFile);
                        var pngOpts2 = new PNGSaveOptions();
                        pngOpts2.compression = 6;
                        tDoc.saveAs(outFile, pngOpts2, true, Extension.LOWERCASE);
                        tDoc.close(SaveOptions.DONOTSAVECHANGES);
                        tiffFile.remove();
                        saved = true;
                    } catch(tifErr) {}
                }

                if (!saved) {
                    writeResult('ERROR:All PNG save methods failed.');
                }

                dupDoc.close(SaveOptions.DONOTSAVECHANGES);

                // Clean up: remove temp alpha channel from original doc, then close it
                try {
                    for (var j = 0; j < doc.channels.length; j++) {
                        if (doc.channels[j].name === 'PPTQuickSelMask') {
                            doc.channels[j].remove();
                            break;
                        }
                    }
                } catch(chErr) {}
                doc.close(SaveOptions.DONOTSAVECHANGES);

                if (saved) {
                    writeResult('' + selL + ',' + selT + ',' + selR + ',' + selB + ',' + docW + ',' + docH);
                }
            }
        }
    }
} catch(globalErr) {
    writeResult('ERROR:' + globalErr.message);
}
"
                .Replace("__CUTOUT__", cutoutPath)
                .Replace("__RESULT__", resultPath);
        }

        /// <summary>
        /// Cleanup JSX: closes the active document without saving (used on cancel).
        /// </summary>
        private static string BuildCleanupJsx()
        {
            return @"
#target photoshop
try {
    if (app.documents.length > 0) {
        app.activeDocument.close(SaveOptions.DONOTSAVECHANGES);
    }
} catch(e) {}
";
        }

        // ═══════════════════════════════════════════════════════════════════
        //  C# Control Panel (shown while user works in Photoshop)
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Builds the TopMost WinForms control-panel dialog. This is NOT modal
        /// to Photoshop — user can freely switch between PS and this dialog.
        /// </summary>
        private static Form BuildControlPanel()
        {
            var form = new Form
            {
                Text = "Quick Select \u2014 Control Panel",
                Size = new Size(440, 330),
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                ShowIcon = false,
                BackColor = Color.White
            };

            var headerLabel = new Label
            {
                Text = "Make Your Selection in Photoshop",
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 102, 204),
                AutoSize = true,
                Location = new Point(20, 18)
            };

            var instructionLabel = new Label
            {
                Text = "1.  Switch to Photoshop (it should be open with your image).\r\n\r\n" +
                       "2.  Use the Quick Selection Tool  (W key)  to paint\r\n" +
                       "     over the area you want to extract.\r\n\r\n" +
                       "3.  Hold  Alt  and paint to subtract from the selection.\r\n\r\n" +
                       "4.  When satisfied, come back here and click  Import.",
                Font = new Font("Segoe UI", 9.5f),
                Location = new Point(20, 55),
                Size = new Size(395, 145)
            };

            var btnImport = new Button
            {
                Text = "\u2714  Import Selection",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Size = new Size(180, 42),
                Location = new Point(50, 220),
                BackColor = Color.FromArgb(0, 120, 212),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnImport.FlatAppearance.BorderSize = 0;
            btnImport.Click += (s, e) => { form.DialogResult = DialogResult.OK; };

            var btnCancel = new Button
            {
                Text = "Cancel",
                Font = new Font("Segoe UI", 10f),
                Size = new Size(120, 42),
                Location = new Point(250, 220),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 1;
            btnCancel.Click += (s, e) => { form.DialogResult = DialogResult.Cancel; };

            form.Controls.AddRange(new Control[] { headerLabel, instructionLabel, btnImport, btnCancel });
            form.AcceptButton = btnImport;
            form.CancelButton = btnCancel;

            return form;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  Result file polling
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Polls for the result file written by the Phase 2 JSX script.
        /// Returns the result string or null on timeout.
        /// </summary>
        private static string WaitForResultFile(string resultFilePath, int timeoutSeconds)
        {
            int elapsed = 0;
            while (elapsed < timeoutSeconds * 1000)
            {
                if (File.Exists(resultFilePath))
                {
                    Thread.Sleep(300); // let PS finish writing
                    try
                    {
                        return File.ReadAllText(resultFilePath).Trim();
                    }
                    catch { /* file locked, retry */ }
                }

                Thread.Sleep(500);
                elapsed += 500;

                // Keep UI responsive
                System.Windows.Forms.Application.DoEvents();
            }
            return null;
        }

        // ═══════════════════════════════════════════════════════════════════
        //  Photoshop EXE locator
        // ═══════════════════════════════════════════════════════════════════

        /// <summary>
        /// Locates the Photoshop executable by checking well-known paths
        /// and the registry App Paths.
        /// </summary>
        private static string FindPhotoshopExe()
        {
            // 1. Check well-known install locations
            string[] knownPaths = new[]
            {
                @"C:\Program Files\Adobe\Adobe Photoshop 2026\Photoshop.exe",
                @"C:\Program Files\Adobe\Adobe Photoshop 2025\Photoshop.exe",
                @"C:\Program Files\Adobe\Adobe Photoshop 2024\Photoshop.exe",
                @"C:\Program Files\Adobe\Adobe Photoshop CC 2023\Photoshop.exe",
                @"C:\Program Files\Adobe\Adobe Photoshop CC 2022\Photoshop.exe",
            };

            foreach (string p in knownPaths)
                if (File.Exists(p)) return p;

            // 2. Search all Adobe Photoshop* folders in Program Files
            try
            {
                string programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                string adobeDir = Path.Combine(programFiles, "Adobe");
                if (Directory.Exists(adobeDir))
                {
                    foreach (string dir in Directory.GetDirectories(adobeDir, "Adobe Photoshop*"))
                    {
                        string candidate = Path.Combine(dir, "Photoshop.exe");
                        if (File.Exists(candidate)) return candidate;
                    }
                }
            }
            catch { }

            // 3. Registry App Paths
            try
            {
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(
                    @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Photoshop.exe"))
                {
                    if (key != null)
                    {
                        string val = key.GetValue(null)?.ToString();
                        if (!string.IsNullOrEmpty(val) && File.Exists(val)) return val;
                    }
                }
            }
            catch { }

            return null;
        }
    }
}
