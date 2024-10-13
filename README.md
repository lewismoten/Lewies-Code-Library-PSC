# Lewies Code Library PSC

This collection showcases open-source projects I published on Planet Source Code (PSC) in the late 1990s and early 2000s. Some scripts can be found on the PSC Index maintained by 

[@Planet-Source-Code](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md).

While many scripts are outdated due to advancements in web development, they represent the cutting edge of technology at the time.

## Highlights

**File Upload Scripts (Classic ASP):** These award-winning scripts enabled file uploads on shared hosting platforms, bypassing vendor restrictions on COM objects costing upwards of $700. Later versions leveraged the ADODB Stream to optimize file parsing, a technique later superseded by ASP.NET's built-in download functionality.

- [Upload Files Without COM](./ASP/UploadFileWithoutCOM)
- [Upload Files Without COM version 2](./ASP/UploadFileWithoutCOMv2)
- [Upload Files Without COM version 3](./ASP/UploadFileWithoutCOMv3)

**SHA-1 Hash Algorithm:** This script, available in Classic ASP and JavaScript, gained unexpected popularity. While included for historical reference, SHA-1 vulnerabilities discovered in 2005 make it unsuitable for modern security applications.

- Classic ASP: [SHA-1 Hash Algorithm](./ASP/SHA-1HashAlgorithm/)
- JavaScript: [Secure Hash Algorithm SHA-1][./JavaScript/SecureHashAlgorithm(SHA-1)/SHA-1]

## Script Organization:

Scripts are categorized by language and date of publication on PSC (July 2000 - August 2004). Some predate my PSC account. In total, the collection comprises 220 scripts, including a few documents.

PSC grouped ASP, VBScript, and VBA as one "world," so server-side and client-side scripts may be mixed within language folders.

- [Classic ASP / VBScript](./ASP/README.md) - *103*
- [ASP.Net](./ASPNet/README.md) - *8*
- [C/C++](./C/README.md) - *1*
- [C#](./CSharp/README.md) - *7*
- [VB.Net](./VBNet/README.md) - *18*
- [Visual Basic](./VB/README.md) - *33*
- [SQL](./SQL/README.md) - *32*
- [JavaScript](./JavaScript/README.md) - *17*
