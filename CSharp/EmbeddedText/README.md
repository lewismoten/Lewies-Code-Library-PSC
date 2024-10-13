# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [C#](../README.md)

### Get Embedded Text Files

*8/5/2004 11:55:55 PM*

Load an embedded text from from within your assembly and assign it to a variable. I primarily use this when storing sql, data, or some pieces of code that I need to have access to. Simply add text files to your solution and change the "Build Action" property to "Embedded Resource" on each one. No need to distribute extra files with your programs - they get compiled as resources inside executables and DLLs.

![Screenshot of Get Embedded Text Files](/screenshot.jpg)



