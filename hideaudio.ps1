# --------------------------------------------------------------------
# Copyright (c) 2020 Hiroaki Kawashima
# This software is released under the MIT License, see LICENSE.txt.
#
# Description:
#     Change settings of audio clip inserted in each slide
#     Input: pptxfile, Output: ppt-out/pptxfile
#
# Usage (use with a batch file):
#     See Usage in *.bat
#
# Usage (run directly):
#     1. Open powershell
#     2. > Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
#     3. > .\hideaudio.ps1 <pptx file>
#     4. Check 'ppt-out' folder
#     5. > Set-ExecutionPolicy -Scope CurrentUser Restricted
# --------------------------------------------------------------------

# Change setting of one slide (xml)
function ChangeAudioSetting ($xmlfile)
{
    Write-Host -NoNewline "[$xmlfile] "
    $xmldoc = [xml](Get-Content -Encoding UTF8 $xmlfile)

    $nsMgr = New-Object System.Xml.XmlNamespaceManager $xmldoc.NameTable  
    $nsMgr.AddNamespace('p', $xmldoc.DocumentElement.GetAttribute('xmlns:p'))
    $nsMgr.AddNamespace('a', $xmldoc.DocumentElement.GetAttribute('xmlns:a'))

    # Audio file exist?
    $checkaudio = $xmldoc.SelectSingleNode('//a:audioFile', $nsMgr)
    if ($checkaudio -eq $null){
        Write-Host 'This slide has no audio'
        return 1
    }

    # Set the first audio as the target
    $node1 = $xmldoc.SelectSingleNode('/p:sld/p:timing/p:tnLst/p:par/p:cTn[@nodeType="tmRoot"]/p:childTnLst/p:seq/p:cTn[@nodeType="mainSeq"]/p:childTnLst/p:par/p:cTn', $nsMgr)
    if ($node1 -eq $null){  # 「開始: クリック時」になっていれば対象外
        Write-Host 'This audio is not a default setting'
        return 1
    }

    # Convert default setting to onBegin
    $childnode1 = $node1.SelectSingleNode('p:stCondLst', $nsMgr)
    if ($childnode1 -eq $null){
        Write-Host 'Cannot find sdContLst'
        return 1
    }
    # Add cond
    $elm = $xmldoc.CreateElement('p:cond', $nsMgr.LookupNamespace('p'))  # namespace必要
    $elm.SetAttribute('evt', 'onBegin')
    $elm.SetAttribute('delay', '0')
    $elm.AppendChild($xmldoc.CreateElement('p:tn', $nsMgr.LookupNamespace('p'))).SetAttribute('val', '2')
    $childnode1.AppendChild($elm) | Out-Null

    $childnode2 = $node1.SelectSingleNode('p:childTnLst/p:par/p:cTn/p:childTnLst/p:par/p:cTn[@presetClass="mediacall" and @nodeType="clickEffect"]', $nsMgr)
    if ($childnode2 -eq $null){  # Skip if already auto begin
        Write-Host 'Cannot find clickEffect'
        return 1
    }
    $childnode2.SetAttribute('nodeType', 'afterEffect')

    # --- Hide audio icon ---
    $node2 = $node1.SelectSingleNode('/p:sld/p:timing/p:tnLst/p:par/p:cTn[@nodeType="tmRoot"]/p:childTnLst/p:audio/p:cMediaNode', $nsMgr)
    if ($node2 -ne $null){
        $node2.SetAttribute('showWhenStopped', '0')
    }else{
        Write-Host 'Cannot find cMediaNode'
        return 1
    }

    $xmldoc.Save($xmlfile)
    Write-Host 'Setting changed!'
    return 0
}


#=====================================================================
# Main: Input is one pptx file
#=====================================================================
if ($args.Length -ne 1) {  # pptx filename
    Write-Host 'Need one argument'
    exit 1
}
$OUTPUT_DIR = 'ppt-out'
$WORK_DIR = '__work__'
$pptfile = $args[0]

New-Item -ItemType Directory -Force -Path $OUTPUT_DIR | Out-Null
New-Item -ItemType Directory -Force -Path $WORK_DIR | Out-Null

Write-Host $pptfile
if (-not (Test-Path $pptfile)){
    Write-Host "No such file: $pptfile"
    exit 1
}
$pptfname = [System.IO.Path]::GetFileName($pptfile)
$pptfbase = [System.IO.Path]::GetFileNameWithoutExtension($pptfile)
$zipfile = "$WORK_DIR\$pptfbase.zip"
$pptworkdir = "$WORK_DIR\$pptfbase"
# Write-Host $zipfile
Copy-Item -Path $pptfile -Destination $zipfile -Force

if (Test-Path $pptworkdir){
    Remove-Item $pptworkdir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $pptworkdir | Out-Null

Expand-Archive -Path $zipfile -DestinationPath $pptworkdir

$slidedir = "$pptworkdir\ppt\slides"
$num_files = (Get-ChildItem -File $slidedir -Filter slide*.xml).Count
Write-Host "Converting $num_files slides"
for ($i=1; $i -le $num_files; $i++){
  $f = "$slidedir\slide$i.xml"
  ChangeAudioSetting $f | Out-Null
}

$newzipfile = "$OUTPUT_DIR\$pptfbase.zip"
$newpptfile = "$OUTPUT_DIR\$pptfname"
Compress-Archive -Path $pptworkdir\* -DestinationPath $newzipfile -Force
Move-Item $newzipfile $newpptfile -Force
Remove-Item $zipfile
Remove-Item -Recurse $pptworkdir
Write-Host "Output: $newpptfile"

exit 0
