# 3rd Party Software Downloader

## Project: 
  - This is part of my MDT automation Project

## What it does:
The script crawls through the 3rd party websites, looking for specific tags in the html and auto navigates to find the download link. Then it will download the files and store them in a folder. Once downloadedm, it will build a CI that can be parsed with another script (working) importing into SCCM or other software delivery 

## Works on:
 - Currently it only these wsf files (https://github.com/PowerShellCrack/MDTDeployApplications) 
 - Only updates products that are download usign this script (https://github.com/PowerShellCrack/Get3rdPartySoftware)


Powershell script to download the latest updates for:
  - Adobe Reader 
  - Adobe Reader Updates (and MUI)
  - Adobe Acrobat Reader DC Updates (and MUI)
  - Flash Player Active X
  - Flash Player NPAPI (Firefox)
  - Flash Player PPAPI (Chrome)
  - Shockwave (full, slim and msi)
  - Google Chrome Enterprise Edition (msi)
  - Google Chrome Standalone (exe)
  - Firefox (x86)
  - Firefox (x64)
  - Notepadd++
  - 7Zip (x64) - MSI and EXE
  - 7Zip (x86) - MSI and EXE
  - VLC Player (x64)
  - VLC Player (x86)
  - Java 8 (x86)
  - Java 8 (x64)
