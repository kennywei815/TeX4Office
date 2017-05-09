TeX4Office
======

[TeX4Office](https://github.com/kennywei815/TeX4Office) is a VBA- and C#-based Microsoft Office add-in which allows users to embed LaTeX codes into Word, Excel, and PowerPoint easily. 

- Support all MS Office products: Word, Excel, and PowerPoint.
- Support MS Office 2003 to Office 2016.
- Mac OS Support is also in massive development.

# Prerequisite
You need to install these packages first:
- [Ghostscript](https://www.ghostscript.com/download/gsdnld.html) for PDF file processing
- A LaTeX distribution: [MikTeX](https://miktex.org/download) (easier to install on Windows, recommended) or [TeX Live](https://www.tug.org/texlive/) (cross-platform)

# Quick Start

### Install
1. `git clone https://github.com/kennywei815/TeX4Office.git`
2. `cd TeX4Office/`
3. `.\setup.bat`

### Usage

If you know how to use LaTeX, it is very easy to use TeX4Office. Select "New/Edit LaTeX display" from the "Add-Ins" tab of the ribbon, and you will get a editor where you can type your equation:

![alt text](https://github.com/kennywei815/TeX4Office/blob/master/Editor.png)

Type any valid LaTeX code, and click on "Output". TeX4Office will compile your code into LaTeX, generate an image from it and insert it into Office.

![alt text](https://github.com/kennywei815/TeX4Office/blob/master/Office_Add-Ins.png)

If you need to change something in the equation, just select the image, then click on "New/Edit LaTeX display" again, and the TeX4Office editor will re-appear so you can edit the LaTeX code.

You can also treat the equation as an ordinary PowerPoint image. For example, it can be grouped, animated, rotated, moved, and resized. Further editing of the equation will preserve all these changes.

When you save the presentation, both the image and the LaTeX code are stored. This means that you can display your presentation on any computer, even computers on which TeX4Office is not installed (no more missing fonts!). Of course, equations can only be edited if you install TeX4Office.


# License
Copyright (c) 2017 Cheng-Kuan Wei Licensed under the Apache License.

This software contains code derived from Jonathan Le Roux and Zvika Ben-Haim's IguanaTeX project which is originally released under the Creative Commons Attribution 3.0 License and combined into TeX4Office as a whole under Apache License 2.0 .