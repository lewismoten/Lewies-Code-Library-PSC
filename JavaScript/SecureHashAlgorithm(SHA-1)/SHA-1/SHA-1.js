/*
Title: The Secure Hash Algorithm (SHA-1)

Category: Computer Security

Description:
	The Secure Hash Algorithm takes a message of less than 264 bits 
	in length and produces a 160-bit message digest which is designed
	so that it should be computationaly expensive to find a text which
	matches a given hash. ie if you have a hash for document A, H(A),
	it is difficult to find a document B which has the same hash, and
	even more difficult to arrange that document B says what you want
	it to say.

Notice:
	The 160-bit message digest is returned as 40 hex characters.
	Each pair represents 1 byte.

Algorithm Credits:
	National Institute of Standards and Technology (NIST)
	FIPS PUB 186
	"Digital Signature Standard"
	U.S. Department of Commerce
	May 1994

Code Port Author: 
	Lewis Moten
	lewis@moten.com
	http://www.lewismoten.com
	April 13, 2001

References:
	Book: 
		Applied Cryptography by Bruce Schneier
		Chapter 18.7 pages 442 to 445
	URL:
		FIPS Publication 180-1
		http://www.itl.nist.gov/fipspubs/fip180-1.htm (html)
		http://csrc.nist.gov/fips/fip180-1.txt (text)
	Visual Basic Version:
		Secure Hash Algorithm SHA-1
		Hosted at http://www.planet-source-code.com
		http://www.planetsourcecode.com/xq/ASP/txtCodeId.13565/lngWId.1/qx/vb/scripts/ShowCode.htm
		Author: Peter Girard
		
For further information on SHA, see: 
	B. Schneier. 
		Applied Cryptography: Protocols, Algorithms, and Source Code in C. 
		Wiley
		2nd Edition
		1995.. 
	B. Preneel. 
		Analysis and Design of Cryptographic Hash Functions. 
		Ph.D. 
		Thesis, Katholieke University Leuven
		1993. 
	M.J.B. Robshaw. 
		MD2, MD4, MD5, SHA and Other Hash Functions.
		Technical Report TR-101
		version 4.0
		RSA Laboratories
		July 1995

*/

// global vars

// Define four constants used in the algorithm
var k = new Array(3)
k[0] = SHA1_Add(0, 0x5a827999)
k[1] = SHA1_Add(0, 0x6ed9eba1)
k[2] = SHA1_Add(0, 0x8f1bbcdc)
k[3] = SHA1_Add(0, 0xca62c1d6)

var StrHexCharacters = "0123456789abcdef" // hex characters

// -----------------------------------------------------------------------------
function SHA1_Hash(pStrMessage)
{
	  
	var W = new Array(80)

	var M = SHA1_MakeBlocks(pStrMessage)

	// Initialize 5 32-bit variables
	var A = SHA1_Add(0, 0x67452301)
	var B = SHA1_Add(0, 0xefcdab89)
	var C = SHA1_Add(0, 0x98badcfe)
	var D = SHA1_Add(0, 0x10325476)
	var E = SHA1_Add(0, 0xc3d2e1f0)

	/*
		first, the five variables are copied into different
		variables: a gets A, b gets B, c gets C, d gets D, 
		and e gets E
	*/

	var a = A
	var b = B
	var c = C
	var d = D
	var e = E
	
	/*
		The main loop processes the message 512 bits at a
		time and continues for as many 512-bit blocks as 
		are in the message.
		
		(16 bytes = 512 bits)
		
	*/
	for(var i = 0; i < M.length; i += 16)
	{
	
		for(var t = 0; t < 80; t++)
		{
			/*
				The message block is transformed from 16 32-bit words
				(M[0] to M[15]) to 90 32-bit words (W[0] to W[79]
				using the following algorithm:
			*/
			if(t < 16)
			{
				W[t] = M[i + t]
			}
			else
			{
				// t = 16 to 79
				W[t] = SHA1_LeftCircularShift(W[t-3] ^ W[t-8] ^ W[t-14] ^ W[t-16], 1)
			}
			
			/*
				each operation performs a nonliniear function of three
				of a, b, c, d, and e and then does shifting and adding
				similar to MD5
			*/
			/*
				As an interesting aside, the original SHA specification
				did not have the left circular shift.  The change
				"corrects a technical flaw that made the standard less
				secure than had been thought".  The NSA has refused to
				elaborate on the exact nature of the flaw.
			*/
			temp = SHA1_Add(SHA1_Add(SHA1_LeftCircularShift(a, 5), SHA_f(t, b, c, d)), SHA1_Add(SHA1_Add(e, W[t]), SHA_K(t)))
			
			e = d
			d = c
			c = SHA1_LeftCircularShift(b, 30)
			b = a
			a = temp
		}

		/*
			After all of this, a, b, c, d, and e are added to A, B, C, D,
			and E respecitively, and the algorithm continues with the next
			block of data.
		*/
		A = SHA1_Add(a, A)
		B = SHA1_Add(b, B)
		C = SHA1_Add(c, C)
		D = SHA1_Add(d, D)
		E = SHA1_Add(e, E)
	}

	// The final output is the concatenation of A, B, C, D, and E
	return(SHA1_Hex(A) + SHA1_Hex(B) + SHA1_Hex(C) + SHA1_Hex(D) + SHA1_Hex(E))
}
// -----------------------------------------------------------------------------
function SHA1_Hex(pLngValue)
{
	var lStrHex = ''
	var lLngPositionIndex
	for(var lLngPositionIndex = 7; lLngPositionIndex >= 0; lLngPositionIndex--)
	{
		lStrHex += StrHexCharacters.charAt((pLngValue >> (lLngPositionIndex * 4)) & 0x0f)
	}
	return(lStrHex)
}
// -----------------------------------------------------------------------------
function SHA1_MakeBlocks(pStrMessage)
{
	var lLngMaxBlock = ((pStrMessage.length + 8) >> 6) + 1
	var lStrBlockAry = new Array(lLngMaxBlock * 16)
	var lLngIndex
	/*
		We convert our message into 16-word blocks
		
		1 word    =  4 bytes =  32 bits
		16 words  = 64 bytes = 512 bits
		
		This will make it easier for us to proces the data in the
		main loop.
	*/
	for(lLngIndex = 0; lLngIndex < lLngMaxBlock * 16; lLngIndex++) 
	{
		lStrBlockAry[lLngIndex] = 0
	}
	
	for(lLngIndex = 0; lLngIndex < pStrMessage.length; lLngIndex++)
	{
		lStrBlockAry[lLngIndex >> 2] |= pStrMessage.charCodeAt(lLngIndex) << (24 - (lLngIndex % 4) * 8)
	}

	/*
		First, the message is padded to make it a multiple of 512
		bits long.  Padding is exactly the same as in MD5: First
		append a one, then as many zeros as necessary to make it
		64 bits short of a multiple of 512, and finally a 64-bit
		representation of the length of the message before padding.
	*/
	
	lStrBlockAry[lLngIndex >> 2] |= 0x80 << (24 - (lLngIndex % 4) * 8)
	lStrBlockAry[lLngMaxBlock * 16 - 1] = pStrMessage.length * 8

	return(lStrBlockAry)
}
// -----------------------------------------------------------------------------
function SHA1_Add(pLngValue1, pLngValue2)
{
	/*
		Add two numbers, but allow wrapping to negative numbers
	*/
	var lLngA = (pLngValue1 & 0xFFFF) + (pLngValue2 & 0xFFFF)
	var lLngB = (pLngValue1 >> 16) + (pLngValue2 >> 16) + (lLngA >> 16)
	var lLngResult = (lLngB << 16) | (lLngA & 0xFFFF)
	return(lLngResult)
}
// -----------------------------------------------------------------------------
function SHA1_LeftCircularShift(pLngValue, pLngBits)
{
	return(pLngValue << pLngBits) | (pLngValue >>> (32 - pLngBits))
}
// -----------------------------------------------------------------------------
function SHA_f(t, x, y, z)
{
	/*
		Each operation performs a nonlinear function of three
		of a, b, c, d, and e, and then does shifting and adding
		similar to MD5
	*/

	// 0 to 19
	if(t < 20) return((x & y) | ((~x) & z))
	// 20 to 39
	if(t < 40) return(x ^ y ^ z)
	// 40 to 59
	if(t < 60) return((x & y) | (x & z) | (y & z))
	// 60 to 79
	return x ^ y ^ z
  
  }
// -----------------------------------------------------------------------------
function SHA_K(t)
{
	// 0 to 19
	if(t < 20) return(k[0])
	// 20 to 39
	if(t < 40) return(k[1])
	// 40 to 59
	if(t < 60) return(k[2])
	// 60 to 79
	return(k[3])
}
// -----------------------------------------------------------------------------