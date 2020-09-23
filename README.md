<div align="center">

## Link URLs


</div>

### Description

Finds any URL found within specified text and creates a hyper link for http, https, ftp, and email addresses.
 
### More Info
 
asContent - Content to be parsed for URLs

This code assumes that you have vbScript 5.0 (or higher) installed on your server. If not, you will instantly receive an error on your page.

returns the content with HTML encoded hyperlinks.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Lewis E\. Moten III](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/lewis-e-moten-iii.md)
**Level**          |Intermediate
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Strings](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/strings__4-26.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/lewis-e-moten-iii-link-urls__4-6298/archive/master.zip)

### API Declarations

Copyright (c) 2000, Lewis Moten. All rights reserved. Send all mofications to the author.


### Source Code

```
Function LinkURLs(ByRef asContent)
	Dim loRegExp	' Regular Expression Object (Requires vbScript 5.0 and above)
	' If no content was received, exit the function
	If asContent = "" Then Exit Function
	' Create Regular Expression object
	Set loRegExp = New RegExp
	' Keep finding links after the first one.
	loRegExp.Global = True
	' Ignore upper/lower case
	loRegExp.IgnoreCase = True
	' Look for URLs
	loRegExp.Pattern = "((ht|f)tps?://\S+[/]?[^\.])([\.]?.*)"
	' Link URLs
	LinkURLs = loRegExp.Replace(asContent, "<A href=""$1"">$1</A>$3")
	' Look for email addresses
	loRegExp.Pattern = "(\S+@\S+.\.\S\S\S?)"
	' Link email addresses
	LinkURLs = loRegExp.Replace(LinkURLs, "<A href=""mailto:$1"">$1</A>")
	' Release regular expression object
	Set oRegExp = Nothing
End Function
```

