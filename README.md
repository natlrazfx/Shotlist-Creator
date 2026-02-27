# ShotlistCreator
This script is a powerful tool for DaVinci Resolve Studio users, specifically designed for those working on storyboards, VFX breakdowns, and notes. It allows you to export marker data and stills directly to an Excel file, helping streamline your workflow and improve communication with vendors, clients, and artists.

### Current Version: 2.1.2
[![Watch the video](https://img.youtube.com/vi/f1luGUWT4PQ/maxresdefault.jpg)](https://youtu.be/f1luGUWT4PQ)

## Compatibility

macOS Ventura, Sonoma with DaVinci Resolve Studio 18, 19  
Windows 10, 11 Pro with DaVinci Resolve Studio 18, 19  
Python 3.10.9 (https://www.python.org/downloads/release/python-3109/)  


## Instruction Manual (for both systems):
There are two ways to use this script:

**1. From the box (pre-compiled executable):**
Unzip the downloaded folder to your preferred location. Inside the folder dist, you will find the executable file ShotlistCreator.  
  [download win pre-compiled executable](https://drive.google.com/file/d/1bj_FhSeDl47fcXVf6rjhiat9ovwBXJRo/view?usp=sharing)    
  [download mac pre-compiled executable](https://drive.google.com/file/d/1bkdBp0rl9xUwwEEEXbDTJ2QrdGBNwaO7/view?usp=sharing)

**2. Running directly from DaVinci Resolve Studio:**  
Copy the file ShotlistCreator.py to the DaVinci Resolve Utility scripts folder:  
*For macOS:* /Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/  
*For Windows:* C:\ProgramData\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\  
Make sure the following Python modules are installed:  
*PySide6*  
*pynput*  
*Pillow*  
*xlsxwriter*  
*DaVinciResolveScript (comes with DaVinci Resolve Studio)*  

## How it works:
  1. Open DaVinci Resolve Studio and load your project.
  2. Go to keyboard customization and assign a key for "Next Marker" (Playback > Next Marker ("0")). This setup is required once. If you run ShotlistCreator.py directly from DaVinci Resolve Studio, you can modify the hotkey in the script and then assign it in the keyboard customization.
  3. Ensure that the album stills1 (in the color page) is empty. This is crucial for the script to function correctly.
  4. Run the script. A dialog box will prompt you to select options such as deleting stills from the album on the color page, setting the timeline timecode, choosing which metadata to extract, and defining the thumbnail size. The script will navigate through the timeline markers, capture thumbnails, and export the marker data and stills to an Excel file in your chosen folder.

[![Watch the video](https://img.youtube.com/vi/lGYmBYw0BuA/maxresdefault.jpg)](https://youtu.be/lGYmBYw0BuA)  

## Additional Info:

1. For annotations, create a paint node in the Fusion page and add your notes there. Marker annotations and burn-in information will not be exported.
2. The exported file is optimized for size, making it easy to convert to PDF or upload to Google Sheets.
3. On macOS, ensure you grant Terminal accessibility access in Privacy settings.
4. On macOS, it’s recommended to launch DaVinci Resolve Studio from Contents-MacOS-Resolve for better performance.
5. This script works only with the Studio version of DaVinci Resolve.

For a Resolve script to be executed from an external folder, the script needs to know of the API location.
You may need to set the these environment variables to allow for your Python installation to pick up the appropriate dependencies as shown below:

    Mac OS X:
    RESOLVE_SCRIPT_API="/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting"
    RESOLVE_SCRIPT_LIB="/Applications/DaVinci Resolve/DaVinci Resolve.app/Contents/Libraries/Fusion/fusionscript.so"
    PYTHONPATH="$PYTHONPATH:$RESOLVE_SCRIPT_API/Modules/"

    Windows:
    RESOLVE_SCRIPT_API="%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting"
    RESOLVE_SCRIPT_LIB="C:\Program Files\Blackmagic Design\DaVinci Resolve\fusionscript.dll"
    PYTHONPATH="%PYTHONPATH%;%RESOLVE_SCRIPT_API%\Modules\"


## Support
If ShotlistCreator helps your workflow, you can support the project here:

[Support ShotlistCreator](https://aescripts.com/shotlist-creator-for-davinci-resolve/)

*Footage is provided by Freepik.*

## Installer Build (macOS + Windows)

This repository now includes automated installer builds for both platforms:

- macOS: `.pkg` installer
- Windows: `.exe` setup installer

### Build trigger

Installers are built automatically by GitHub Actions when you push a tag like:

```bash
git tag -a v2.1.2 -m "Release v2.1.2"
git push origin v2.1.2
```

You can also run the workflow manually from the Actions tab (`Build Installers`).

### Output artifacts

After workflow completion, download artifacts:

- `ShotlistCreator-macOS-<version>.pkg`
- `ShotlistCreator-<version>-Windows-Setup.exe`

On tag builds, these files are also attached directly to the GitHub Release page.

### Local build helpers

- macOS: `./packaging/macos/build_pkg.sh 2.1.2`
- Windows (PowerShell): `./packaging/windows/build_installer.ps1 -Version 2.1.2`

## End-User Install Experience

For non-technical users, distribution should be:

1. User downloads installer from GitHub Release:
`ShotlistCreator-<version>-Windows-Setup.exe` or `ShotlistCreator-<version>-macOS.pkg`.
2. User installs normally (Next/Install flow).
3. User opens DaVinci Resolve Studio and then launches ShotlistCreator.

Important requirement:
- DaVinci Resolve Studio must already be installed on the user's machine.

If Resolve is missing/not running, the app now shows a clear startup message instead of a Python traceback.
