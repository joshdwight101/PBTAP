# **PBTAP (Powershell Based Twain Acquisition Program)**

**Created by Joshua Dwight** *Available on GitHub: [joshdwight101/PBTAP](https://www.google.com/search?q=https://github.com/joshdwight101/PBTAP)*

PBTAP is a hybrid PowerShell and C\# utility designed for high-performance TWAIN image acquisition, tailored specifically for healthcare and dental imaging environments. It features a lightweight WinForms GUI with asynchronous scanning to prevent UI lockups and built-in image manipulation tools.

## **Features**

* **Native TWAIN Integration:** Leverages NTwain v3 to interface directly with scanners and dental imaging hardware.  
* **Optimized for Healthcare:** Forces 16-bit grayscale capture by default, ideal for digital X-Rays and clinical imaging.  
* **Non-Blocking UI:** Scans execute asynchronously in a background C\# thread.  
* **On-the-Fly Image Processing:** Includes custom C\#-based matrix manipulation for instantaneous Brightness, Contrast, and Sharpen adjustments.  
* **Persistent Settings:** Automatically remembers your last used TWAIN device and export format.  
* **Comparison View:** Real-time "Compare" button to toggle between original scan and diagnostic adjustments.

## **Requirements**

* Windows OS with PowerShell 5.1+  
* .NET Framework (System.Windows.Forms, System.Drawing)  
* NTwain.dll (v3 build) placed in the same directory as the script.

## **Installation & Setup**

1. Clone or download the repository to your local machine.  
2. Ensure you have the NTwain.dll file (net462 version recommended).  
3. Place NTwain.dll in the same directory as PBTAP.ps1.  
4. Right-click PBTAP.ps1 and select **Run with PowerShell**.

## **Usage**

1. **Acquire Image:** Click the blue button to initialize your scanner and capture the image.  
2. **Adjust:** Use the Brightness and Contrast sliders to clarify details. Use the Sharpen slider to enhance edge definition.  
3. **Compare:** Hold the "Compare" button in the top-right of the viewer to see the raw original image.  
4. **Apply:** Click "Apply Changes" to finalize your edits.  
5. **Save:** Choose your format (JPG, PNG, etc.) and save the image to disk.