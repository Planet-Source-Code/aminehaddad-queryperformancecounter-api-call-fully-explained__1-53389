<div align="center">

## QueryPerformanceCounter API Call Fully Explained


</div>

### Description

Most people use GetTickCount, but to be honest this makes GetTickCount look like a joke. QueryPerformanceCounter finds the time in microseconds, and not milliseconds (like GetTickCount). Why use this? If you wanted to see the time that it took for the cpu to proccess an API call, GetTickCount would find 0, because it takes less then a millisecond. But QueryPerformanceCounter would find the real time. I suggest using this because it is much more precise.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2004-04-25 17:06:32
**By**             |[AmineHaddad](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aminehaddad.md)
**Level**          |Advanced
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[QueryPerfo1737624252004\.zip](https://github.com/Planet-Source-Code/aminehaddad-queryperformancecounter-api-call-fully-explained__1-53389/archive/master.zip)

### API Declarations

```
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
```





