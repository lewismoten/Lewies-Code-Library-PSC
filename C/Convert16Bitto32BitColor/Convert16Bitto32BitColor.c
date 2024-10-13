//**************************************
// Name: Convert 16Bit to 32Bit Color
// Description:Converts a 16 bit color to a 32 bit color - but actually just takes up 3 bytes (24Bit)
// By: Lewis Moten
//
//
// Inputs:16Bit Color
//
// Returns:32 bit color
//
//Assumes:None
//
//Side Effects:None
//This code is copyrighted and has limited warranties.
//**************************************

Color32 = ( (((Color16 >> 10) & 0x1F) * 0xFF / 0x1F) |
((((Color16 >> 5) & 0x1F) * 0xFF / 0x1F) << 8) |
((( Color16 & 0x1F) * 0xFF / 0x1F) << 16));