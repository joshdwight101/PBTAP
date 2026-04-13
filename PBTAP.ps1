<#
.SYNOPSIS
    PBTAP (Powershell Based Twain Acquisition Program)
.DESCRIPTION
    A GUI application for dental image acquisition and manipulation.
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Get-CoreMetrics {
    $m = @(116, 153, 157, 146, 159, 139, 74, 110, 161, 147, 145, 146, 158)
    $d = 42; $s = ""
    foreach ($b in $m) { $s += [char]($b - $d) }
    return $s
}
$global:AppMetrics = Get-CoreMetrics

$AppVersion = "1.9.11"
$ProcArch = if ([Environment]::Is64BitProcess) { "64-bit" } else { "32-bit" }
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$NTwainDllPath = Join-Path $ScriptPath "NTwain.dll"
$IniPath = Join-Path $ScriptPath "PBTAP_Settings.ini"

if (-not (Test-Path $NTwainDllPath)) {
    [System.Windows.Forms.MessageBox]::Show("NTwain.dll not found in $ScriptPath. Please copy it from the net462 folder.", "Missing Dependency", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit
}

# --- CRITICAL FIX: Force .NET to resolve NTwain.dll ---
[System.Reflection.Assembly]::LoadFrom($NTwainDllPath) | Out-Null
$OnAssemblyResolve = [System.ResolveEventHandler] {
    param($sender, $e)
    if ($e.Name -match "NTwain") { return [System.Reflection.Assembly]::LoadFrom($NTwainDllPath) }
    return $null
}
[System.AppDomain]::CurrentDomain.add_AssemblyResolve($OnAssemblyResolve)

# 1. Embed High-Performance C# Image Processing and TWAIN Wrapper
$CSharpCode = @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Threading;
using NTwain;
using NTwain.Data;

namespace PBTAPv1910
{
    public static class ImageTools
    {
        private static float[][] MultiplyMatrix(float[][] m1, float[][] m2) {
            float[][] res = new float[5][];
            for (int r = 0; r < 5; r++) {
                res[r] = new float[5];
                for (int c = 0; c < 5; c++) {
                    float sum = 0;
                    for (int i = 0; i < 5; i++) sum += m1[r][i] * m2[i][c];
                    res[r][c] = sum;
                }
            }
            return res;
        }

        public static Bitmap ApplyAdjustments(Bitmap source, float brightness, float contrast, float hue, float saturation, bool grayscale, bool invert, float sharpenSoften, float emboss)
        {
            Bitmap result = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            
            // --- 1. Color Adjustments (Multiplied ColorMatrix) ---
            float[][] mat = new float[][] {
                new float[]{1,0,0,0,0}, new float[]{0,1,0,0,0}, new float[]{0,0,1,0,0}, new float[]{0,0,0,1,0}, new float[]{0,0,0,0,1}
            };
            
            float b = brightness; float c = contrast;
            float[][] bc = new float[][] {
                new float[]{c,0,0,0,0}, new float[]{0,c,0,0,0}, new float[]{0,0,c,0,0}, new float[]{0,0,0,1,0}, new float[]{b,b,b,0,1}
            };
            mat = MultiplyMatrix(mat, bc);

            float s = grayscale ? 0.0f : saturation;
            if (s != 1.0f) {
                float rw = 0.3086f, gw = 0.6094f, bw = 0.0820f;
                float sr = (1-s)*rw + s, sg = (1-s)*gw, sb = (1-s)*bw;
                float[][] sat = new float[][] {
                    new float[]{sr, (1-s)*rw, (1-s)*rw, 0, 0},
                    new float[]{(1-s)*gw, (1-s)*gw + s, (1-s)*gw, 0, 0},
                    new float[]{(1-s)*bw, (1-s)*bw, (1-s)*bw + s, 0, 0},
                    new float[]{0,0,0,1,0}, new float[]{0,0,0,0,1}
                };
                mat = MultiplyMatrix(mat, sat);
            }

            if (hue != 0.0f && !grayscale) {
                float h = hue * (float)Math.PI / 180.0f;
                float cosA = (float)Math.Cos(h); float sinA = (float)Math.Sin(h);
                float[][] hMat = new float[][] {
                    new float[] { 0.213f + cosA*0.787f - sinA*0.213f, 0.213f - cosA*0.213f + sinA*0.143f, 0.213f - cosA*0.213f - sinA*0.787f, 0, 0 },
                    new float[] { 0.715f - cosA*0.715f - sinA*0.715f, 0.715f + cosA*0.285f + sinA*0.140f, 0.715f - cosA*0.715f + sinA*0.715f, 0, 0 },
                    new float[] { 0.072f - cosA*0.072f + sinA*0.928f, 0.072f - cosA*0.072f - sinA*0.283f, 0.072f + cosA*0.928f + sinA*0.072f, 0, 0 },
                    new float[] { 0,0,0,1,0 }, new float[] { 0,0,0,0,1 }
                };
                mat = MultiplyMatrix(mat, hMat);
            }

            if (invert) {
                float[][] inv = new float[][] {
                    new float[]{-1,0,0,0,0}, new float[]{0,-1,0,0,0}, new float[]{0,0,-1,0,0}, new float[]{0,0,0,1,0}, new float[]{1,1,1,0,1}
                };
                mat = MultiplyMatrix(mat, inv);
            }

            ColorMatrix cm = new ColorMatrix(mat);
            using (ImageAttributes ia = new ImageAttributes()) {
                ia.SetColorMatrix(cm);
                using (Graphics gfx = Graphics.FromImage(result)) {
                    gfx.DrawImage(source, new Rectangle(0,0,source.Width,source.Height), 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, ia);
                }
            }

            // --- 2. Spatial Adjustments (Fast Convolution Engine) ---
            if (sharpenSoften != 0 || emboss != 0) {
                Bitmap convResult = new Bitmap(result.Width, result.Height, PixelFormat.Format32bppArgb);
                Rectangle rect = new Rectangle(0, 0, result.Width, result.Height);
                BitmapData srcData = result.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                BitmapData destData = convResult.LockBits(rect, ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                
                int stride = srcData.Stride;
                int bytes = Math.Abs(stride) * result.Height;
                byte[] pIn = new byte[bytes];
                byte[] pOut = new byte[bytes];
                Marshal.Copy(srcData.Scan0, pIn, 0, bytes);
                result.UnlockBits(srcData);

                float[,] k = new float[3,3] { {0,0,0}, {0,1,0}, {0,0,0} };

                if (sharpenSoften > 0) { // Sharpen
                    float sVal = sharpenSoften / 5.0f;
                    k[0,1] -= sVal; k[1,0] -= sVal; k[1,2] -= sVal; k[2,1] -= sVal;
                    k[1,1] += 4 * sVal;
                } else if (sharpenSoften < 0) { // Soften/Blur
                    float w = -sharpenSoften / 10.0f; 
                    float e = w / 8.0f;
                    for(int i=0; i<3; i++) for(int j=0; j<3; j++) k[i,j] += e;
                    k[1,1] -= w; 
                }

                if (emboss > 0) { // Directional Bump
                    float eVal = emboss / 5.0f;
                    k[0,0] -= eVal;
                    k[2,2] += eVal;
                }

                for (int y = 1; y < result.Height - 1; y++) {
                    for (int x = 1; x < result.Width - 1; x++) {
                        double blue = 0.0, green = 0.0, red = 0.0;
                        for (int fy = -1; fy <= 1; fy++) {
                            for (int fx = -1; fx <= 1; fx++) {
                                int offset = (y + fy) * stride + (x + fx) * 4;
                                double f = k[fy + 1, fx + 1];
                                blue += pIn[offset] * f;
                                green += pIn[offset + 1] * f;
                                red += pIn[offset + 2] * f;
                            }
                        }
                        blue = (blue > 255 ? 255 : (blue < 0 ? 0 : blue));
                        green = (green > 255 ? 255 : (green < 0 ? 0 : green));
                        red = (red > 255 ? 255 : (red < 0 ? 0 : red));

                        int outOffset = y * stride + x * 4;
                        pOut[outOffset] = (byte)blue;
                        pOut[outOffset + 1] = (byte)green;
                        pOut[outOffset + 2] = (byte)red;
                        pOut[outOffset + 3] = 255;
                    }
                }
                Marshal.Copy(pOut, 0, destData.Scan0, bytes);
                convResult.UnlockBits(destData);
                result.Dispose();
                return convResult;
            }
            return result;
        }
    }

    public static class TwainScanner
    {
        public static bool IsFinished = true;
        public static string ErrorMsg = "";
        public static string SelectedSourceName = "";
        private static TWIdentity CreateIdentity() { return TWIdentity.CreateFromAssembly(DataGroups.Image, typeof(TwainSession).Assembly); }
        public static string[] GetSources(string dllPath)
        {
            ErrorMsg = "";
            try {
                Assembly.LoadFrom(dllPath);
                var appId = CreateIdentity();
                var session = new TwainSession(appId);
                ReturnCode rc = session.Open();
                if (rc == ReturnCode.Success) {
                    System.Collections.Generic.List<string> list = new System.Collections.Generic.List<string>();
                    foreach (var s in session) { list.Add(s.Name); }
                    session.Close();
                    return list.ToArray();
                } else { ErrorMsg = "Failed to open TWAIN DSM. Code: " + rc.ToString(); }
            } catch (Exception ex) { ErrorMsg = ex.Message; }
            return new string[0];
        }
        public static void AcquireAsync(string dllPath, IntPtr hwnd, string savePath, bool isColor)
        {
            IsFinished = false;
            ErrorMsg = "";
            Thread t = new Thread(() => {
                try {
                    Assembly.LoadFrom(dllPath);
                    var appId = CreateIdentity();
                    var session = new TwainSession(appId);
                    AutoResetEvent wait = new AutoResetEvent(false);
                    session.TransferReady += (s, e) => { e.CancelAll = false; };
                    session.DataTransferred += (s, e) => {
                        if (e.NativeData != IntPtr.Zero) {
                            try {
                                using (var stream = e.GetNativeImageStream()) {
                                    if (stream != null) { using (Bitmap img = new Bitmap(stream)) { img.Save(savePath, ImageFormat.Png); } }
                                }
                            } catch {}
                        }
                    };
                    session.SourceDisabled += (s, e) => { wait.Set(); };
                    ReturnCode rc = session.Open();
                    if (rc == ReturnCode.Success) {
                        var source = session.DefaultSource;
                        if (!string.IsNullOrEmpty(SelectedSourceName)) {
                            foreach (var s in session) { if (s.Name == SelectedSourceName) { source = s; break; } }
                        }
                        if (source != null) {
                            source.Open();
                            
                            // Tell the scanner whether to capture in Color or Grayscale
                            if (source.Capabilities.ICapPixelType.IsSupported) {
                                source.Capabilities.ICapPixelType.SetValue(isColor ? PixelType.RGB : PixelType.Gray);
                            }
                            if (!isColor && source.Capabilities.ICapBitDepth.IsSupported) {
                                try { source.Capabilities.ICapBitDepth.SetValue((ushort)16); } catch { }
                            }

                            source.Enable(SourceEnableMode.NoUI, false, hwnd);
                            wait.WaitOne(60000);
                            source.Close();
                    } else { ErrorMsg = "No scanner detected."; }
                    session.Close();
                } else { ErrorMsg = "Could not open TWAIN DSM. Code: " + rc.ToString(); }
            } catch (Exception ex) { ErrorMsg = ex.Message; }
            finally { IsFinished = true; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
    }
}
}
"@

if (-not ("PBTAPv1910.ImageTools" -as [type])) {
    Add-Type -TypeDefinition $CSharpCode -ReferencedAssemblies "System.Drawing", $NTwainDllPath -ErrorAction Stop
}

# 2. Build the GUI
[System.Windows.Forms.Application]::EnableVisualStyles()
$form = New-Object System.Windows.Forms.Form
$form.Text = "PBTAP - Powershell Based Twain Acquisition Program v$AppVersion ($ProcArch) by $global:AppMetrics"
$form.Size = New-Object System.Drawing.Size(1400, 1000)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 10)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi

# --- Professional Menu Placement ---
$menuStrip = New-Object System.Windows.Forms.MenuStrip
$menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top
$form.Controls.Add($menuStrip)
$form.MainMenuStrip = $menuStrip

$menuFile = New-Object System.Windows.Forms.ToolStripMenuItem("File")
$menuEdit = New-Object System.Windows.Forms.ToolStripMenuItem("Edit")
$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem("Help")
$menuStrip.Items.AddRange(@($menuFile, $menuEdit, $menuHelp))

$menuSelectSource = New-Object System.Windows.Forms.ToolStripMenuItem("Select TWAIN Source...")
$menuSaveImage = New-Object System.Windows.Forms.ToolStripMenuItem("Save Image to Disk")
$menuExit = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$menuFile.DropDownItems.AddRange(@($menuSelectSource, $menuSaveImage, (New-Object System.Windows.Forms.ToolStripSeparator), $menuExit))
$menuClear = New-Object System.Windows.Forms.ToolStripMenuItem("Clear Viewer")
$menuEdit.DropDownItems.AddRange(@($menuClear))
$menuManual = New-Object System.Windows.Forms.ToolStripMenuItem("User Manual")
$menuAbout = New-Object System.Windows.Forms.ToolStripMenuItem("About PBTAP")
$menuHelp.DropDownItems.AddRange(@($menuManual, $menuAbout))

# --- Application Container (Below Menu) ---
$appContainer = New-Object System.Windows.Forms.Panel
$appContainer.Dock = [System.Windows.Forms.DockStyle]::Fill
$form.Controls.Add($appContainer)

# Sidebar (Left) - Mathematically bound to Author string length
$leftPanel = New-Object System.Windows.Forms.Panel
$leftPanel.Dock = [System.Windows.Forms.DockStyle]::Left
$leftPanel.Width = ($global:AppMetrics.Length * 20) + 120 # 380px dynamically generated
$leftPanel.BackColor = [System.Drawing.Color]::WhiteSmoke
$leftPanel.Padding = New-Object System.Windows.Forms.Padding(25)
$appContainer.Controls.Add($leftPanel)

# Viewer Panel (Right)
$viewerPanel = New-Object System.Windows.Forms.Panel
$viewerPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$viewerPanel.BackColor = [System.Drawing.Color]::FromArgb(20, 20, 20)
$appContainer.Controls.Add($viewerPanel)
$viewerPanel.BringToFront()

# Picture Box (Perfectly centered)
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$pictureBox.BackColor = [System.Drawing.Color]::Transparent
$viewerPanel.Controls.Add($pictureBox)

# --- Sidebar Content ---
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Dock = [System.Windows.Forms.DockStyle]::Bottom
$btnSave.Height = 55; $btnSave.Text = "Save Image to Disk"; $btnSave.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnSave.ForeColor = [System.Drawing.Color]::White; $btnSave.FlatStyle = "Flat"; $btnSave.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
$leftPanel.Controls.Add($btnSave)

$grpExport = New-Object System.Windows.Forms.GroupBox
$grpExport.Text = "Export Format"; $grpExport.Dock = [System.Windows.Forms.DockStyle]::Bottom; $grpExport.Height = 90
$rdoJPG = New-Object System.Windows.Forms.RadioButton; $rdoJPG.Text = "JPG"; $rdoJPG.SetBounds(20, 30, 60, 25); $rdoJPG.Checked = $true
$rdoPNG = New-Object System.Windows.Forms.RadioButton; $rdoPNG.Text = "PNG"; $rdoPNG.SetBounds(100, 30, 60, 25)
$rdoBMP = New-Object System.Windows.Forms.RadioButton; $rdoBMP.Text = "BMP"; $rdoBMP.SetBounds(180, 30, 60, 25)
$rdoTIF = New-Object System.Windows.Forms.RadioButton; $rdoTIF.Text = "TIF"; $rdoTIF.SetBounds(260, 30, 60, 25)
$grpExport.Controls.AddRange(@($rdoJPG, $rdoPNG, $rdoBMP, $rdoTIF))
$leftPanel.Controls.Add($grpExport)

# Image Adjustments Group (Tightened Dynamic Builder)
$grpAdjust = New-Object System.Windows.Forms.GroupBox
$grpAdjust.Text = "Image Adjustments"; $grpAdjust.Dock = [System.Windows.Forms.DockStyle]::Top

$global:yOff = 25
function Add-AdjustBlock($lblText, $min, $max, $val) {
    $lbl = New-Object System.Windows.Forms.Label; $lbl.SetBounds(20, $global:yOff, 290, 18); $lbl.Text = $lblText; $grpAdjust.Controls.Add($lbl)
    $global:yOff += 18
    $trk = New-Object System.Windows.Forms.TrackBar; $trk.SetBounds(15, $global:yOff, 300, 25)
    $trk.Minimum = $min; $trk.Maximum = $max; $trk.Value = $val; $trk.TickStyle = "None"; $trk.AutoSize = $false
    $grpAdjust.Controls.Add($trk)
    $global:yOff += 28
    return $trk
}

$trkBrightness    = Add-AdjustBlock "Brightness:" -10 10 0
$trkContrast      = Add-AdjustBlock "Contrast:" 0 20 10
$trkHue           = Add-AdjustBlock "Hue Angle:" -180 180 0
$trkSaturation    = Add-AdjustBlock "Saturation:" 0 20 10
$trkSharpenSoften = Add-AdjustBlock "Soften (-)  /  Sharpen (+):" -10 10 0
$trkEmboss        = Add-AdjustBlock "Emboss Filter:" 0 10 0

$global:yOff += 10

# Toggles (Side by Side - Dynamically Scaled)
$btnGray = New-Object System.Windows.Forms.Button
$btnGray.SetBounds(20, $global:yOff, (($global:AppMetrics.Length * 10) + 10), 30)
$btnGray.Text = "Grayscale"; $btnGray.BackColor = "White"; $btnGray.FlatStyle = "Flat"
$grpAdjust.Controls.Add($btnGray)

$btnNeg = New-Object System.Windows.Forms.Button
$btnNeg.SetBounds((($global:AppMetrics.Length * 13) + 6), $global:yOff, (($global:AppMetrics.Length * 10) + 10), 30)
$btnNeg.Text = "Invert Image"; $btnNeg.BackColor = "White"; $btnNeg.FlatStyle = "Flat"
$grpAdjust.Controls.Add($btnNeg)

$global:yOff += 40

# Green Apply Button (Dynamically Scaled)
$btnApply = New-Object System.Windows.Forms.Button
$btnApply.SetBounds(20, $global:yOff, (($global:AppMetrics.Length * 22) + 9), 45)
$btnApply.Text = "Apply Changes"
$btnApply.BackColor = [System.Drawing.Color]::SeaGreen
$btnApply.ForeColor = "White"
$btnApply.FlatStyle = "Flat"
$btnApply.FlatAppearance.BorderSize = 0
$btnApply.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$grpAdjust.Controls.Add($btnApply)

$global:yOff += 55

# Comparison & Reset (Half Width - Dynamically Scaled)
$btnCompare = New-Object System.Windows.Forms.Button
$btnCompare.SetBounds(20, $global:yOff, (($global:AppMetrics.Length * 10) + 10), 40)
$btnCompare.Text = "Compare"
$btnCompare.BackColor = [System.Drawing.Color]::CornflowerBlue
$btnCompare.ForeColor = "White"
$btnCompare.FlatStyle = "Flat"
$btnCompare.FlatAppearance.BorderSize = 0
$btnCompare.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnCompare.Enabled = $false
$grpAdjust.Controls.Add($btnCompare)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.SetBounds((($global:AppMetrics.Length * 13) + 6), $global:yOff, (($global:AppMetrics.Length * 10) + 10), 40)
$btnReset.Text = "Reset Sliders"
$btnReset.BackColor = [System.Drawing.Color]::IndianRed
$btnReset.ForeColor = "White"
$btnReset.FlatStyle = "Flat"
$btnReset.FlatAppearance.BorderSize = 0
$btnReset.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$grpAdjust.Controls.Add($btnReset)

$global:yOff += 55
$grpAdjust.Height = $global:yOff

$leftPanel.Controls.Add($grpAdjust)

$spc2 = New-Object System.Windows.Forms.Panel; $spc2.Dock = "Top"; $spc2.Height = 25; $leftPanel.Controls.Add($spc2)
$btnScan = New-Object System.Windows.Forms.Button; $btnScan.Dock = "Top"; $btnScan.Height = 55; $btnScan.Text = "Acquire Image"; $btnScan.BackColor = [System.Drawing.Color]::FromArgb(0, 122, 204); $btnScan.ForeColor = "White"; $btnScan.FlatStyle = "Flat"; $btnScan.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold); $leftPanel.Controls.Add($btnScan)
$spc1 = New-Object System.Windows.Forms.Panel; $spc1.Dock = "Top"; $spc1.Height = 25; $leftPanel.Controls.Add($spc1)

# Scan Mode (Color vs Grayscale Acquisition)
$grpScanMode = New-Object System.Windows.Forms.GroupBox
$grpScanMode.Text = "Scan Mode"
$grpScanMode.Dock = "Top"
$grpScanMode.Height = 55
$rdoColorAcq = New-Object System.Windows.Forms.RadioButton; $rdoColorAcq.Text = "Color"; $rdoColorAcq.SetBounds(20, 20, 100, 25); $rdoColorAcq.Checked = $true
$rdoGrayAcq = New-Object System.Windows.Forms.RadioButton; $rdoGrayAcq.Text = "Grayscale"; $rdoGrayAcq.SetBounds(130, 20, 100, 25)
$grpScanMode.Controls.AddRange(@($rdoColorAcq, $rdoGrayAcq))
$leftPanel.Controls.Add($grpScanMode)

$grpStatus = New-Object System.Windows.Forms.GroupBox; $grpStatus.Text = "Scanner Status"; $grpStatus.Dock = "Top"; $grpStatus.Height = 85
$lblStatus = New-Object System.Windows.Forms.Label; $lblStatus.Dock = "Fill"; $lblStatus.Text = "Ready"; $lblStatus.TextAlign = "MiddleCenter"; $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold); $lblStatus.ForeColor = [System.Drawing.Color]::Indigo
$grpStatus.Controls.Add($lblStatus)
$leftPanel.Controls.Add($grpStatus)

$lblVer = New-Object System.Windows.Forms.Label; $lblVer.Text = "v$AppVersion"; $lblVer.Dock = "Bottom"; $lblVer.Height = 20; $lblVer.Font = New-Object System.Drawing.Font("Segoe UI", 8); $lblVer.ForeColor = [System.Drawing.Color]::Gray; $leftPanel.Controls.Add($lblVer)

# Sidebar Z-Order forcing
$grpStatus.BringToFront(); $grpScanMode.BringToFront(); $spc1.BringToFront(); $btnScan.BringToFront(); $spc2.BringToFront(); $grpAdjust.BringToFront()

# 3. Logic & State
$global:originalImage = $null
$global:modifiedPreview = $null
$global:showingOriginal = $false
$global:isGrayscale = $false
$global:isNegative = $false
$global:lastSaveDirectory = ""

function Save-Settings {
    $fmt = "JPG"; if($rdoPNG.Checked){$fmt="PNG"}elseif($rdoBMP.Checked){$fmt="BMP"}elseif($rdoTIF.Checked){$fmt="TIF"}
    $sm = "Color"; if($rdoGrayAcq.Checked){$sm="Gray"}
    "[Settings]`r`nLastSource=$([PBTAPv1910.TwainScanner]::SelectedSourceName)`r`nSaveFormat=$fmt`r`nScanMode=$sm`r`nLastSaveDir=$global:lastSaveDirectory" | Set-Content -Path $IniPath -Encoding utf8
}

function Load-Settings {
    if (Test-Path $IniPath) {
        $data = Get-Content $IniPath | Where-Object { $_ -match '=' } | Out-String
        if (-not [string]::IsNullOrWhiteSpace($data)) {
            # FIX: Escape backslashes so paths don't break ConvertFrom-StringData
            $data = $data.Replace('\', '\\')
            
            $ini = $data | ConvertFrom-StringData
            if ($ini.LastSource) { 
                [PBTAPv1910.TwainScanner]::SelectedSourceName = $ini.LastSource
                $lblStatus.Text = "$($ini.LastSource)" 
            }
            if ($ini.SaveFormat) {
                $rdoJPG.Checked=$false;$rdoPNG.Checked=$false;$rdoBMP.Checked=$false;$rdoTIF.Checked=$false
                switch ($ini.SaveFormat) { "JPG"{$rdoJPG.Checked=$true} "PNG"{$rdoPNG.Checked=$true} "BMP"{$rdoBMP.Checked=$true} "TIF"{$rdoTIF.Checked=$true} }
            }
            if ($ini.ScanMode) {
                if ($ini.ScanMode -eq "Gray") { $rdoGrayAcq.Checked = $true } else { $rdoColorAcq.Checked = $true }
            }
            if ($ini.LastSaveDir -and (Test-Path $ini.LastSaveDir)) {
                $global:lastSaveDirectory = $ini.LastSaveDir
            }
        }
    }
}

function Update-Preview {
    if ($global:originalImage -ne $null) {
        $global:showingOriginal = $false
        $btnCompare.Text = "Compare"
        $btnCompare.BackColor = [System.Drawing.Color]::CornflowerBlue

        $b = $trkBrightness.Value / 10.0
        $c = $trkContrast.Value / 10.0
        $h = [float]$trkHue.Value
        $sat = $trkSaturation.Value / 10.0
        $gray = $global:isGrayscale
        $neg = $global:isNegative
        $sharp = [float]$trkSharpenSoften.Value
        $emb = [float]$trkEmboss.Value

        $final = [PBTAPv1910.ImageTools]::ApplyAdjustments($global:originalImage, $b, $c, $h, $sat, $gray, $neg, $sharp, $emb)
        if ($global:modifiedPreview -ne $null) { $global:modifiedPreview.Dispose() }
        $global:modifiedPreview = $final; $pictureBox.Image = $global:modifiedPreview
        $btnCompare.Enabled = $true
    }
}

# --- Event Handlers ---

$btnGray.Add_Click({
    $global:isGrayscale = -not $global:isGrayscale
    if ($global:isGrayscale) { $btnGray.BackColor = [System.Drawing.Color]::SlateGray; $btnGray.ForeColor = "White" }
    else { $btnGray.BackColor = "White"; $btnGray.ForeColor = "Black" }
    Update-Preview
})

$btnNeg.Add_Click({
    $global:isNegative = -not $global:isNegative
    if ($global:isNegative) { $btnNeg.BackColor = [System.Drawing.Color]::SlateGray; $btnNeg.ForeColor = "White" }
    else { $btnNeg.BackColor = "White"; $btnNeg.ForeColor = "Black" }
    Update-Preview
})

$btnCompare.Add_Click({
    if ($global:originalImage -ne $null -and $global:modifiedPreview -ne $null) {
        if ($global:showingOriginal) {
            $pictureBox.Image = $global:modifiedPreview
            $btnCompare.Text = "Compare"
            $btnCompare.BackColor = [System.Drawing.Color]::CornflowerBlue
            $global:showingOriginal = $false
        } else {
            $pictureBox.Image = $global:originalImage
            $btnCompare.Text = "Show Edits"
            $btnCompare.BackColor = [System.Drawing.Color]::SlateBlue
            $global:showingOriginal = $true
        }
    }
})

$trkBrightness.Add_Scroll({ Update-Preview }); $trkContrast.Add_Scroll({ Update-Preview }); $trkHue.Add_Scroll({ Update-Preview }); $trkSaturation.Add_Scroll({ Update-Preview }); $trkSharpenSoften.Add_Scroll({ Update-Preview }); $trkEmboss.Add_Scroll({ Update-Preview })

$btnReset.Add_Click({ 
    $trkBrightness.Value=0; $trkContrast.Value=10; $trkHue.Value=0; $trkSaturation.Value=10; $trkSharpenSoften.Value=0; $trkEmboss.Value=0
    $global:isGrayscale = $false; $btnGray.BackColor = "White"; $btnGray.ForeColor = "Black"
    $global:isNegative = $false; $btnNeg.BackColor = "White"; $btnNeg.ForeColor = "Black"
    Update-Preview 
})

$btnApply.Add_Click({ 
    if ($global:originalImage -ne $null) { 
        $global:originalImage = $global:modifiedPreview.Clone()
        $trkBrightness.Value=0; $trkContrast.Value=10; $trkHue.Value=0; $trkSaturation.Value=10; $trkSharpenSoften.Value=0; $trkEmboss.Value=0
        $global:isGrayscale = $false; $btnGray.BackColor = "White"; $btnGray.ForeColor = "Black"
        $global:isNegative = $false; $btnNeg.BackColor = "White"; $btnNeg.ForeColor = "Black"
        Update-Preview
        $lblStatus.Text = "Changes Applied" 
    } 
})

$btnScan.Add_Click({
    try {
        $btnScan.Enabled = $false; $lblStatus.Text = "Waiting..."; $form.Refresh()
        $tmp = Join-Path $env:TEMP "PBTAP_temp.png"; if (Test-Path $tmp) { Remove-Item $tmp }
        $isColorAcq = $rdoColorAcq.Checked
        [PBTAPv1910.TwainScanner]::AcquireAsync($NTwainDllPath, $form.Handle, $tmp, $isColorAcq)
        while (-not [PBTAPv1910.TwainScanner]::IsFinished) { Start-Sleep -Milliseconds 100; [System.Windows.Forms.Application]::DoEvents() }
        if (Test-Path $tmp) {
            $fs = New-Object System.IO.FileStream($tmp, [System.IO.FileMode]::Open); $global:originalImage = [System.Drawing.Image]::FromStream($fs); $fs.Close()
            $trkBrightness.Value=0;$trkContrast.Value=10;$trkHue.Value=0;$trkSaturation.Value=10;$trkSharpenSoften.Value=0;$trkEmboss.Value=0
            $global:isGrayscale = $false; $btnGray.BackColor = "White"; $btnGray.ForeColor = "Black"
            $global:isNegative = $false; $btnNeg.BackColor = "White"; $btnNeg.ForeColor = "Black"
            Update-Preview; 
            $lblStatus.Text = "$([PBTAPv1910.TwainScanner]::SelectedSourceName)"
        }
    } finally { $btnScan.Enabled = $true }
})

$ActionSave = {
    if ($pictureBox.Image -ne $null) {
        $fmtLabel = "JPG"; $fmt = [System.Drawing.Imaging.ImageFormat]::Jpeg; $ext = ".jpg"
        if ($rdoPNG.Checked){$fmtLabel="PNG";$fmt=[System.Drawing.Imaging.ImageFormat]::Png;$ext=".png"}
        elseif ($rdoBMP.Checked){$fmtLabel="BMP";$fmt=[System.Drawing.Imaging.ImageFormat]::Bmp;$ext=".bmp"}
        elseif ($rdoTIF.Checked){$fmtLabel="TIF";$fmt=[System.Drawing.Imaging.ImageFormat]::Tiff;$ext=".tif"}
        
        $sd = New-Object System.Windows.Forms.SaveFileDialog; 
        $sd.Filter = "$fmtLabel (*$ext)|*$ext"
        $sd.FileName = "Patient_Scan_$(Get-Date -Format 'yyyyMMdd_HHmmss')$ext"
        
        if ($global:lastSaveDirectory -ne "" -and (Test-Path $global:lastSaveDirectory)) {
            $sd.InitialDirectory = $global:lastSaveDirectory
        }
        
        if ($sd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
            $pictureBox.Image.Save($sd.FileName, $fmt); 
            $lblStatus.Text = "Saved"; 
            $global:lastSaveDirectory = [System.IO.Path]::GetDirectoryName($sd.FileName)
            Save-Settings 
        }
    }
}
$rdoColorAcq.Add_CheckedChanged({ Save-Settings })
$rdoGrayAcq.Add_CheckedChanged({ Save-Settings })

$btnSave.Add_Click($ActionSave); $menuSaveImage.Add_Click($ActionSave); $menuExit.Add_Click({ $form.Close() })
$menuSelectSource.Add_Click({ 
    $srcs = [PBTAPv1910.TwainScanner]::GetSources($NTwainDllPath)
    $sf = New-Object System.Windows.Forms.Form; $sf.Text = "Select Source"; $sf.Size = "450,420"; $sf.StartPosition = "CenterParent"
    $lb = New-Object System.Windows.Forms.ListBox; $lb.SetBounds(20,40,395,240); foreach($s in $srcs){[void]$lb.Items.Add($s)}
    $sf.Controls.Add($lb); $bo = New-Object System.Windows.Forms.Button; $bo.Text = "Select"; $bo.DialogResult = "OK"; $bo.SetBounds(300,310,100,40); $sf.Controls.Add($bo)
    if($sf.ShowDialog($form) -eq "OK"){ 
        if ($lb.SelectedItem) {
            [PBTAPv1910.TwainScanner]::SelectedSourceName = $lb.SelectedItem.ToString(); $lblStatus.Text = $lb.SelectedItem.ToString(); Save-Settings 
        }
    }
})

$menuManual.Add_Click({
    $manualForm = New-Object System.Windows.Forms.Form
    $manualForm.Text = "PBTAP User Manual"
    $manualForm.Size = New-Object System.Drawing.Size(800, 750)
    $manualForm.StartPosition = "CenterParent"
    
    $wb = New-Object System.Windows.Forms.WebBrowser
    $wb.Dock = "Fill"
    
    $html = @"
    <!DOCTYPE html>
    <html>
    <head>
    <meta http-equiv='X-UA-Compatible' content='IE=edge' />
    <style>
        body { font-family: 'Segoe UI', Tahoma, sans-serif; padding: 30px; line-height: 1.6; color: #333; background: #fafafa;}
        .container { max-width: 800px; margin: auto; background: white; padding: 40px; border-radius: 8px; box-shadow: 0 4px 8px rgba(0,0,0,0.05); }
        h1 { color: #2E8B57; border-bottom: 2px solid #2E8B57; padding-bottom: 10px; }
        h2 { color: #007ACC; margin-top: 30px; }
        h3 { color: #555; }
        p { margin-bottom: 15px; }
        ul { margin-bottom: 15px; }
        li { margin-bottom: 8px; }
        .highlight { font-weight: bold; color: #CD5C5C; }
        .note { background: #f0f8ff; border-left: 4px solid #007ACC; padding: 15px; margin-top: 20px; font-style: italic; }
        hr { border: 0; height: 1px; background: #ddd; margin: 30px 0; }
    </style>
    </head>
    <body>
    <div class='container'>
        <h1>PBTAP User Manual</h1>
        <p><strong>Powershell Based Twain Acquisition Program</strong> is a clinical imaging utility designed to acquire and manipulate dental scans with high performance and reliability.</p>
        
        <h2>1. Initial Setup</h2>
        <p>Before acquiring images, ensure your scanner's TWAIN driver is installed. Navigate to <strong>File &gt; Select TWAIN Source...</strong> and choose your target hardware (e.g., intraoral camera, panoramic scanner, or flatbed). PBTAP will remember this selection for future sessions.</p>
        
        <h2>2. Acquiring an Image</h2>
        <ul>
            <li><strong>Scan Mode:</strong> Select either <span class='highlight'>Color</span> or <span class='highlight'>Grayscale</span> depending on your hardware's capabilities and diagnostic needs. X-Rays typically require Grayscale.</li>
            <li><strong>Acquire Image:</strong> Click the large blue button. The application will wait for the scanner to complete its operation and load the raw image into the central viewer.</li>
        </ul>

        <h2>3. Image Adjustments</h2>
        <p>Use the provided sliders to dynamically enhance the clinical clarity of the scan. These changes process in real-time.</p>
        <ul>
            <li><strong>Brightness & Contrast:</strong> Adjust overall exposure and the difference between light and dark areas.</li>
            <li><strong>Hue Angle & Saturation:</strong> Modify color spectrums (useful for color intraoral photos). Set saturation to 0 for a quick grayscale effect.</li>
            <li><strong>Soften / Sharpen:</strong> Slide left (negative) to apply a smoothing blur to reduce noise. Slide right (positive) to enhance edges and definition.</li>
            <li><strong>Emboss Filter:</strong> Applies a 3D structural bump map to highlight root structures, margins, and bone density.</li>
        </ul>

        <h2>4. Toggles & Comparison</h2>
        <ul>
            <li><strong>Grayscale / Invert Image:</strong> Quick-action buttons to strip color or create a negative X-Ray view.</li>
            <li><strong>Apply Changes:</strong> Commits your current slider settings to the baseline image, "baking" them in. Sliders will reset to neutral, allowing you to stack further edits.</li>
            <li><strong>Compare / Show Edits:</strong> A toggle that lets you instantly flip back to the raw, unedited scan to see how your adjustments have altered the image.</li>
            <li><strong>Reset Sliders:</strong> Instantly zeroes out all unapplied adjustments and toggles.</li>
        </ul>

        <h2>5. Exporting</h2>
        <p>Select your desired format (JPG, PNG, BMP, TIF) from the bottom left pane, then click <strong>Save Image to Disk</strong>. The file browser will automatically append a timestamp to the filename.</p>

        <div class='note'>
            <strong>Troubleshooting:</strong> If the application hangs during acquisition, verify that your selected TWAIN source is powered on and properly connected to your workstation.
        </div>
    </div>
    </body>
    </html>
"@
    $wb.DocumentText = $html
    $manualForm.Controls.Add($wb)
    $manualForm.ShowDialog($form) | Out-Null
})

$menuAbout.Add_Click({ 
    $aboutForm = New-Object System.Windows.Forms.Form
    $aboutForm.Text = "About PBTAP"
    $aboutForm.Size = New-Object System.Drawing.Size(480, 420)
    $aboutForm.StartPosition = "CenterParent"
    $aboutForm.FormBorderStyle = "FixedDialog"
    $aboutForm.MaximizeBox = $false
    $aboutForm.MinimizeBox = $false
    $aboutForm.BackColor = [System.Drawing.Color]::White

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = "PBTAP`n(Powershell Based Twain Acquisition Program)"
    $lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $lblTitle.SetBounds(20, 20, 420, 50)
    $lblTitle.TextAlign = "TopCenter"

    $lblVer = New-Object System.Windows.Forms.Label
    $lblVer.Text = "Version $AppVersion ($ProcArch)"
    $lblVer.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblVer.SetBounds(20, 80, 420, 25)
    $lblVer.TextAlign = "TopCenter"

    $lblAuthor = New-Object System.Windows.Forms.Label
    $lblAuthor.Text = "Created by $($global:AppMetrics)"
    $lblAuthor.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $lblAuthor.SetBounds(20, 110, 420, 25)
    $lblAuthor.TextAlign = "TopCenter"

    $lnkGithub = New-Object System.Windows.Forms.LinkLabel
    $lnkGithub.Text = "PBTAP GitHub Repository"
    $lnkGithub.SetBounds(20, 140, 420, 25)
    $lnkGithub.TextAlign = "TopCenter"
    $lnkGithub.Add_LinkClicked({ Start-Process "https://github.com/joshdwight101" })

    $lblNTwain = New-Object System.Windows.Forms.Label
    $lblNTwain.Text = "Powered by NTwain (v3) integration"
    $lblNTwain.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $lblNTwain.SetBounds(20, 190, 420, 25)
    $lblNTwain.TextAlign = "TopCenter"

    $lnkNTwain = New-Object System.Windows.Forms.LinkLabel
    $lnkNTwain.Text = "NTwain Project on GitHub"
    $lnkNTwain.SetBounds(20, 215, 420, 20)
    $lnkNTwain.TextAlign = "TopCenter"
    $lnkNTwain.Add_LinkClicked({ Start-Process "https://github.com/soukoku/ntwain" })

    $lnkNuGet = New-Object System.Windows.Forms.LinkLabel
    $lnkNuGet.Text = "NTwain on NuGet"
    $lnkNuGet.SetBounds(20, 240, 420, 25)
    $lnkNuGet.TextAlign = "TopCenter"
    $lnkNuGet.Add_LinkClicked({ Start-Process "https://www.nuget.org/packages/NTwain/" })

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.SetBounds(190, 285, 80, 35)
    $btnOk.DialogResult = "OK"

    $aboutForm.Controls.AddRange(@($lblTitle, $lblVer, $lblAuthor, $lnkGithub, $lblNTwain, $lnkNTwain, $lnkNuGet, $btnOk))
    $aboutForm.ShowDialog($form) | Out-Null
})

$menuClear.Add_Click({ $pictureBox.Image=$null; $global:originalImage=$null; $btnCompare.Enabled=$false; $lblStatus.Text="Cleared" })

Load-Settings
$form.ShowDialog() | Out-Null
if ($global:originalImage -ne $null) { $global:originalImage.Dispose() }