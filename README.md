# **PBTAP (Powershell Based Twain Acquisition Program)**

**PBTAP** is a lightweight, high-performance clinical imaging utility designed to acquire and manipulate dental and medical scans directly within a Windows environment. Built with a clean, professional graphical user interface (GUI) via PowerShell and deeply embedded C\# integration, PBTAP offers seamless communication with TWAIN-compatible devices such as intraoral cameras, panoramic x-ray scanners, and standard flatbed scanners.

Developed by **Joshua Dwight**.

## **📖 About the Program**

PBTAP was created to streamline the digital radiography and clinical photography workflow. Often, enterprise imaging software can be bulky, slow, or difficult to quickly adjust on the fly. PBTAP acts as a rapid, standalone acquisition bridge. It allows clinical staff to immediately trigger a scan, dynamically adjust the visual clarity of the resulting image using hardware-accelerated image processing, and export the file safely to the network or local disk.

Because it runs via PowerShell, it requires no complex installation wizards or registry modifications—simply place the script and its dependency in a folder and run it. Under the hood, PBTAP dynamically compiles highly optimized C\# classes directly into memory, delivering the speed of a native desktop application with the ultimate portability of a script.

## **✨ Key Features & Under the Hood**

### **⚙️ Dynamic Architecture Support (x86 & 64-bit)**

PBTAP is completely architecture-agnostic. Whether you launch the script using a 32-bit (x86) or 64-bit (x64) PowerShell host, the application will dynamically detect the runtime environment and adjust its underlying memory allocation. This ensures maximum compatibility with both legacy 32-bit dental drivers and modern 64-bit TWAIN systems without requiring separate downloads.

### **📡 Advanced TWAIN Acquisition**

* **Source Selection:** Automatically detects and lists all installed TWAIN drivers on the workstation.  
* **Dual Scan Modes:** Choose between **Color** (for intraoral photography) or 16-bit **Grayscale** (standard for digital X-Rays) before acquiring.

### **🧠 Smart Session Memory & Autosave**

PBTAP is designed for rapid, successive workflows. It features a built-in memory system (PBTAP\_Settings.ini) that securely stores your preferences so the application is always ready exactly as you left it.

* **Persistent Hardware:** Remembers your last-used scanner and Color/Grayscale scan mode between sessions.  
* **Save Directory Memory:** Automatically defaults to the last network drive or local folder you saved an image to.  
* **Persistent Export Formats:** Remembers your preferred image format (JPG, PNG, BMP, or TIF).  
* **Auto-Timestamping:** Automatically generates a safe filename (e.g., Patient\_Scan\_YYYYMMDD\_HHMMSS) to completely prevent accidental file overwriting or data loss during busy clinical hours.

### **🎛️ Clinical Image Adjustments**

PBTAP utilizes a custom-built, single-pass ColorMatrix multiplier and spatial convolution engine, allowing for real-time edits without stuttering or UI lag.

* **Brightness & Contrast:** Dynamically adjust the exposure and clarity of the raw scan.  
* **Hue & Saturation:** Fine-tune color spectrums for intraoral photos. Slide Saturation to 0 for an instant grayscale effect.  
* **Soften / Sharpen Slider:** A unified spatial convolution slider. Slide left (negative) to apply a smoothing box-blur to reduce grain/noise. Slide right (positive) to enhance and sharpen clinical edges.  
* **Emboss Filter:** Applies a 3D structural bump-map overlay. Highly useful for defining hard edges like roots, margins, or bone structure without destroying diagnostic luminance.  
* **Grayscale & Invert Toggles:** Dedicated push-buttons to instantly strip color or create a true "negative" X-Ray view.

### **🔄 Workflow & Usability**

* **Real-Time Compare:** A one-click toggle to instantly flip between your active edits and the raw, unedited scan to track your changes.  
* **Edit Stacking:** The **Apply Changes** button "bakes" your current adjustments into the image and resets the sliders to zero, allowing you to stack multiple layers of adjustments.  
* **Responsive UI:** The user interface features dynamic, mathematically bound scaling to ensure buttons and sliders never overlap, regardless of Windows display scaling (DPI) settings.

## **🚀 Installation & Usage**

PBTAP requires no traditional installation.

1. Download the latest release from this repository.  
2. Ensure both files are in the same folder:  
   * PBTAP.ps1 (The main application script)  
   * NTwain.dll (The required bridging library)

### **Execution Policy & Deployment**

By default, Windows restricts the execution of unverified PowerShell scripts.

**For Standard/Standalone Use:**

To bypass the default execution policy for a single run, you can launch the application by opening a command prompt or creating a desktop shortcut with the following target:

powershell.exe \-ExecutionPolicy Bypass \-WindowStyle Hidden \-File "C:\\Path\\To\\PBTAP.ps1"

**For Corporate & Clinical Enterprise Environments:**

Running scripts with ExecutionPolicy Bypass is generally discouraged in secured enterprise domains. IT Administrators deploying PBTAP across a clinic should:

1. Digitally sign PBTAP.ps1 using an internal, trusted enterprise Code Signing Certificate.  
2. Distribute the signed script alongside NTwain.dll to the target workstations.  
3. Configure the environment via Group Policy Object (GPO) to allow AllSigned execution or add the certificate publisher to the Trusted Publishers store.

## **🤝 Acknowledgments & Dependencies**

PBTAP’s ability to seamlessly communicate with external scanning hardware is made possible by the incredible **NTwain** library.

Respectable mention and full credit for the underlying TWAIN wrapper go to its original author and contributors. Without this robust open-source library, building a lightweight PowerShell TWAIN application would not be possible.

**NTwain (v3)** \> **Author:** Yin-Chun Wang (soukoku) and NTwain contributors

**GitHub Repository:** [https://github.com/soukoku/ntwain](https://github.com/soukoku/ntwain)

**NuGet Package:** [https://www.nuget.org/packages/NTwain/](https://www.nuget.org/packages/NTwain/)

*Licensed under the MIT License.*

## **📜 License**

This project is licensed under the **MIT License**. See the LICENSE.txt file for full details.

You are free to use this application personally, internally within your clinic, or commercially.

*Disclaimer: This software is provided "as is", without warranty of any kind. It is intended as an acquisition and manipulation utility and should be used in accordance with your local clinical and diagnostic regulations.*