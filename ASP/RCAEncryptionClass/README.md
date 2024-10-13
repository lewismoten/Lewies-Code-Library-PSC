# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [Classic ASP / vbScript](../README.md)

### RCA Encryption Class

*5/9/2001 5:33:51 PM*

RCA is used for most standard public key/private key encryption techniques. Due to its burden on resources, RCA is generally used to encrypt random session keys wich are then used to encode messages with RC4. RCA also brings digital certificates into the picture. This, and many other "public" algorithms using RC4 do not truly make the message secure. Encryption is performed in 8 bit blocks - or 1 byte. Typing "AAAAAA", you will see a pattern begin for form with the encrypted data. If someone were to look at the results from encrypting each letter of the alphabet, your data would be compromized. Also, the public/private keys are limited between a range of 1 and about 65000. This is primarily a starting point if you want to enable a more secure encryption scheme, or you just want to learn more about it.

![Screenshot of RCA Encryption Class](/screenshot.gif)



