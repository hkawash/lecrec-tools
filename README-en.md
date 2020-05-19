[(Japanese)](README.md)

# Compressing audio data in PowerPoint slide files

The bitrate of audio files recorded with PowerPoint seems too large for "speech". This tool tries to shrink the file size of pptx/ppsx files by making lower the bitrate of internal audio data (e.g., inserted by **Insert > Audio > Record Audio**).

- Compress internal m4a audio files to 64kbps bitrate.
- Does not compress images or videos.

## Installation

1. Create a main folder (for example, `lecrec-tools`), a subfolder `ppt-in`, and download the following script to the main folder (use right-click > Save Link As...).
   - Windows: [compress_pptaudio-win.bat](https://github.com/hkawash/lecrec-tools/raw/master/compress_pptaudio-win.bat)
   - macOS/Linux: [compress_pptaudio-mac.sh](https://github.com/hkawash/lecrec-tools/raw/master/compress_pptaudio-mac.sh)
   - For a simpler step, you can download [the archive of this project](https://github.com/hkawash/lecrec-tools/archive/master.zip), and extract the zip file.
2. Download ffmpeg from [this site (zeranoe)](https://ffmpeg.zeranoe.com/builds/), and extract the zip file.
   - Version: Select release build (4.2.2)
   - Architecture: Select your OS
   - Linking: Static
   - You can also download from [ffmpeg original site](https://www.ffmpeg.org/download.html)
3. Copy ffmpeg.exe (or ffmpeg for macOS) from the extracted folder to the main folder created (or extracted) in step 1.

If you already have ffmpeg installed, you can skip step 2 and 3. If have not set the environment variable (path) for ffmpeg yet, you can edit the script (the line with `PATH`) or simply follow step 3 (copy the ffmpeg executable file).

## Usage

1. Put pptx and/or ppsx files under the `ppt-in` folder. (Backup the files, just in case.)
2. Run the script file (<a href="#note1">See also Note</a>).
   - Windows: double click `compress_pptaudio-win.bat`
   - macOS (Linux, Windowsのbash): run `compress_pptaudio-mac.sh` from terminal
3. Output compressed files can be found in the `ppt-out` folder.

<a name="note1"></a>

### Note

#### Windows

- **If you are using Windows 10, you will see a warning message for the first time. Click "More info" and then select "Run anyway".**
- Wait until the opened (black) window is automatically closed.
- Do not run the script again before the window is closed.

#### macOS (Linux, Windows bash)

1. Open bash (terminal). If you use macOS, type "terminal" from Launchpad or Spotlight search.
1. Change the current folder by typing the following command.
    ```
    $ cd /Users/<user name>/Documents/lecrec-tools-master
    ```
   (Press `return` after you typed. `<user name>` should be replaced by your user name of the PC.)
1. Run the script by typing one of the following two commands.
    ```
    $ chmod 755 compress_pptaudio-mac.sh
    $ ./compress_pptaudio-mac.sh
    ```
    (`chmod` is required only once after the download)

### Else

- This script converts internal m4a audiofiles to bitrate 64kbps. If you prefer to use other bitrate, change `BITRATE` inside the script.
- If you use other audio formats, edit the lines that contain `m4a` inside the script.
