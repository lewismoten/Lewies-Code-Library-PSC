# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [Classic ASP / vbScript](../README.md)

### Dynamic Curve Images

*4/29/2001 2:33:00 PM*

The script is able to dynamically create GIF images. The 4 positions of a curve - Top Right/Left and Bottom Right/Left. It also creates a transparent image if the position isn't defined. The images returned are 16x16 pixels - except for the transparent image that is 1x1. The curves are made by two colors (Background and Foreground) that you specify. The script calculates all colors in between and returns an image that appears Anti-Aliased (no jagged edges). This script was made to help users customize there own colors, yet still be able to use images throughout the site that do not lock specific colors into the scheme. This demonstration shows the dynamic image generation in action. 4 rounded corners are created with aliasing to prevent that "blocky" edge on the corners. Simply enter an HTML color (in hex format) in the form below. If an invalid color is specified, then that color will default to either white (Color1) or Black (Color2).

![Screenshot of Dynamic Curve Images](/screenshot.gif)



