# WIM Background Updater by UTZ v1.6

## Description
WIM Background Updater by UTZ is a powerful yet lightweight utility for modifying Windows Imaging Format (WIM) files. Designed for both IT professionals and enthusiasts, this tool enables you to easily update the background image in both Windows Setup and Windows PE environments.

## Key Features
- **Easy-to-Use GUI:** Intuitive graphical interface for selecting the target WIM file, choosing between built-in or custom images, and specifying the image index.
- **Flexible Update Types:** Supports two update modes:
  - **Setup Background Update:** Updates the Windows Setup background with a `.bmp` image.
  - **PE Background Update:** Replaces the Windows PE background using a `.jpg` image.
- **Robust Read-Only Handling:** Automatically detects if the WIM file or its directory is read-only, then copies the file to a writable location and removes the read-only attribute for seamless processing.
- **Automated Logging:** Detailed logging of all operations is saved in `WIM_Updater.log` for troubleshooting and audit purposes.
- **Retry Mechanism:** Automatically retries with a default image index if the specified index is not found.
- **Compatibility:** Built using AutoIt, designed to work on Windows platforms with minimal dependencies.

## Prerequisites
- Windows (tested on Windows 10 & Windows 11)
- AutoIt v3 (runtime and SciTE for editing)
- Administrator privileges (required for processing system files)

## How It Works
1. **Select a WIM File:** Use the file browser to locate and select the WIM file to be updated.
2. **Choose the Image:** Decide whether to use a built-in image or a custom image from your system.
3. **Set the Image Index:** Enter the target image index. If the specified index is not found, the utility automatically retries using a default index.
4. **Process and Update:** The tool extracts necessary files into a temporary folder, updates the WIM file using wimlib-imagex, and then cleans up temporary files.
5. **Logging and Notifications:** Detailed progress and log messages are displayed during processing.

## Installation & Usage
1. Download and extract the release zip file.
2. Run the executable as an administrator.
3. Follow the on-screen instructions to select your WIM file, image, and update type.
4. Monitor the progress via the status window and log file.
5. Upon successful update, the modified WIM file is ready for use.
