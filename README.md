# 3rdPartySoftware
06122018 - Updated to v2


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
  
  
# How
 The script crawls through the 3rd party websites, looking for specific tags in the html and auto navigates to find the download link. Then it will download the files and store them in a folder. Once downloadedm, it will build a CI that can be parsed with another script (working) importing into SCCM or other software delivery 
