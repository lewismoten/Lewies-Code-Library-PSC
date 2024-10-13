# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [SQL](../README.md)

### OutlineSort '3.2.10.2'

*12/3/2002 4:03:40 PM*

This function pads each number found within each decimal to allow sorting on outline numbers. Fixes problems where 2 comes after 10 in string sorts. An example of how this function is used would appear like this: SELECT OutlineNumber FROM Outline ORDER BY dbo.OutlineSort(OutlineNumber)


