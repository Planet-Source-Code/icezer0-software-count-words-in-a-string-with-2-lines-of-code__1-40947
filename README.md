<div align="center">

## Count words in a string with 2 lines of code


</div>

### Description

Have you been using the instr function, or 3rd party functions to find the word count of a block of text. Now you find it in ONLY 2 LINES OF CODE - NO API CALLS - NO MODULES / CLASSES / CONTROLS - PURE VB CODE
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[IceZer0 Software](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/icezer0-software.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/icezer0-software-count-words-in-a-string-with-2-lines-of-code__1-40947/archive/master.zip)





### Source Code

```
Public Function WordCount(Text as String) as Long
'Declare an Array
Dim splitText() as String
splitText = Split(Text, " ")
'Split it using a space charachter for where to split it.
WordCount = Ubound(splitText) + 1
'Set the function return to the end of the array. Ex: See Spot Run! would make splitText(0) = See, splitText(1) = spot, and splitText(2) = Run! - 2 is the highest so 1 is added to it to make the word count
'Good luck on your vb programming - IceZero
End Function
```

