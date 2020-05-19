#!/bin/bash

# ----------------------------------------------------------------------
# Copyright (c) 2020 Hiroaki Kawashima
# This software is released under the MIT License, see LICENSE.txt.
#
# Preparation:
#     1. Download ffmpeg archive and extract the archive file
#     2. Move ffmpeg executable file in the "bin" folder to this folder
#        (which contains this script). 
#           Instead of moving the executable file, you can also set the PATH
#        by setting system environment variable or editting the
#        following code (the line starts with "set PATH").
#       
# Usage:
#     1. Put pptx/ppsx files to "ppt-in" folder
#     2. Run this script (double-click or run from command terminal)
#     3. Compressed pptx/ppsx files can be found in "ppt-out" folder
# ----------------------------------------------------------------------

# PATH="$PATH":"/mnt/c/opt/ffmpeg-4.2.2-win64-static/bin"  # PATHを設定するならここを編集

FFMPEG=./ffmpeg  # Windows なら .exe を付ける

# Check commands
if ! type $FFMPEG > /dev/null 2>&1; then
    echo "Error: Cannot find ffmpeg"
    exit 1
fi

if ! type zip > /dev/null 2>&1; then
    echo "Error: Install [zip] before using this script"
    exit 1
fi
if ! type unzip > /dev/null 2>&1; then
    echo "Error: Install [unzip] before using this script"
    exit 1
fi


# ===========================================================
# Set bitrate here. (Do not insert space around '=')
BITRATE=64k
# ===========================================================

INPUT_DIR=ppt-in
OUTPUT_DIR=ppt-out
WORK_DIR=__work__

if [ ! -d $INPUT_DIR ]; then
    echo "Error: Folder [$INPUT_DIR] does not exist"
    exit 1
fi

ls $INPUT_DIR/*.pp[ts]x
if [ $? -ne 0 ]; then
    echo "Error: No pptx/ppsx files in [$INPUT_DIR] folder"
    exit 1
fi

if [ ! -d $WORK_DIR ]; then
    mkdir -p $WORK_DIR
fi
if [ ! -d $OUTPUT_DIR ]; then
    mkdir -p $OUTPUT_DIR
fi

for f in $INPUT_DIR/*.pp[ts]x; do
    # Setup filenames
    echo "$f"
    pptfname="${f##*/}"
    pptfbase="${pptfname%.*}"
    zipfname=$WORK_DIR/$pptfbase.zip
    pptworkdir=$WORK_DIR/$pptfbase
    echo "$zipfname"
    echo "$pptworkdir"
    cp "$f" "$zipfname"

    # Create folder for pptx content
    if [ -d "$pptworkdir" ]; then
      rm -rf "$pptworkdir"
    fi
    mkdir "$pptworkdir"

    # Expand zip file
    unzip "$zipfname" -d "$pptworkdir"
    if [ ! -d "$pptworkdir"/ppt/media ]; then
        echo "Error: No media folder in pptx/ppsx"
        exit 1
    fi

    # Compress audio files
    for a in "$pptworkdir"/ppt/media/*.m4a; do
        m4afname="${a##*/}"
        $FFMPEG -i "$a" -ab $BITRATE "$WORK_DIR/$m4afname"
        mv "$WORK_DIR/$m4afname" "$a"
    done

    # Archive again
    (cd "$pptworkdir" && zip -9 -r "../../$OUTPUT_DIR/$pptfbase.zip" *)
    mv "$OUTPUT_DIR/$pptfbase.zip" "$OUTPUT_DIR/$pptfname"

    # Remove temoporary files/folders
    rm "$zipfname"
    rm -rf "$pptworkdir"
done
