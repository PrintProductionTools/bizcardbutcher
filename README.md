# BizCardButcher

This is a simple utility which takes one-up print-ready PDF files of business cards and:
* Rotates them to a horizontal layout, if necessary
* Crops them down to an appropriate size
* Imposes them 10-up in a 8.5x11 template suitable for slitting on a Duplo card slitter

By default, the utility is configured to:
* Process all PDF files in a folder named `input` in its own root directory
* Export processed 10-up files in a folder named `output` in its own root directory

These folders can be changed by editing the `sourcePath` and `destinationPath` keys in `BizCardButcher.vshost.exe.config`.

This utility is powered by iTextSharp (https://github.com/itext/itextsharp)
