<div align="center">

## ExtractFileName


</div>

### Description

It extracts a filename from a filepath.
 
### More Info
 
A string comtaining a valide path & filename

Put this function in a module or the declaration section of a form in which this function is needed.

Returns only the filename

It only works with VB6, this of the function 'StrReverse' which is not in previous versions of VB


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[D\. de Haas](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/d-de-haas.md)
**Level**          |Advanced
**User Rating**    |4.3 (170 globes from 40 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/d-de-haas-extractfilename__1-5174/archive/master.zip)





### Source Code

```
Public Function ExtractFileName(ByVal strPath As String) As String
 ' StrReverse is only working in VB6
 strPath = StrReverse(strPath)
 strPath = Left(strPath, InStr(strPath, "\") - 1)
 ExtractFileName = StrReverse(strPath)
End Function
```

