<div align="center">

## Wildcard Matcher


</div>

### Description

This code matches a given string against a pattern which may contain the well known wildcards * and ?, whereas * represents any number of characters, including none, and ? represents any single character. You may want to paste the code into a class module :)
 
### More Info
 
Name as String, Pattern as String. Name is the String you want to match against Pattern.

MatchCase() is the case-sensitive worker function. It makes usage of PreparePattern() to escape the characters # and [ (read Side effects).

Match() is just a wrapper for MatchCase() which does a LCase() on both Name and Pattern parameters.

Returns Boolean True if String matches for Pattern, else returns Boolean False.

This code is written specifically for usage in an IRC bot, thus does escape certain chars which are recognized by the "Like" operator, which this code builds on. Namely, # and [ are taken literally rather than being interpreted as wildcards.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[gridrun](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/gridrun.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/gridrun-wildcard-matcher__1-28751/archive/master.zip)





### Source Code

```
Public Function Match(Name As String, Pattern As String) As Boolean
  If MatchCase(LCase(Name), LCase(Pattern)) Then Match = True
End Function
Public Function MatchCase(Name As String, Pattern As String) As Boolean
  Pattern = PreparePattern(Pattern)
  If Name Like Pattern Then MatchCase = True
End Function
Private Function PreparePattern(Pattern As String) As String
  Pattern = Replace(Pattern, "[", "[[]")
  Pattern = Replace(Pattern, "#", "[#]")
  PreparePattern = Pattern
End Function
```

