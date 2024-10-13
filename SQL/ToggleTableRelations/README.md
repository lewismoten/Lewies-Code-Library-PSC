# [Lewie's Code Library PSC](../../README.md)

Open source projects that I had published to Planet Source Code.

## [SQL](../README.md)

### Toggle Table Relations

*5/29/2002 9:43:53 AM*

Useful for disabling constraints during database replication. When copying data to another database, use this procedure on the target database first. When finnished, run it again to enable the constraints. I used this because I was working on a database with tables referencing themselves. I couldn't add records when previouse records did not exist. This solved the problem.


