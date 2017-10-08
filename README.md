# QXlsx

> *Read this in other languages: [English](README.md), [한국어](README.ko.md)*

![](image/XLSX-64.png)
![](image/qt-64.png)

## Excel reader/writer library

* QXlsx is excel file(*.xlsx) reader/writer library. 
* It is written for static library using Qt 5.

## Test environment
* Qt 5.9.1 (Visual Studio 2017/Windows 64bit) 
* Qt 5.9.1 (Visual Studio 2017/Windows 32bit) 
* Qt 5.9.1 (MingW/Windows 32bit) 
* Qt 5.9.1 (Ubuntu 16/Linux x64) 
* Qt 5.5.1 (MingW/Windows 32bit)
* Qt 5.6.0 (MingW/Windows 32bit) 
* Qt 5.5.0 (Ubuntu 17/Linux i686)
* Qt 5.2.0 (Ubuntu 14/Linux x64)
* Qt 5.0.1 (MingW/Windows 32bit) 

## Hello excel (writing excel file)
```cpp
QXlsx::Document xlsx;
xlsx.write("A1", "Hello Qt!");
xlsx.saveAs("Test.xlsx");
```
## How to build
* see https://github.com/j2doll/QXlsx/wiki

## License
* GPL License https://www.gnu.org/licenses/gpl-3.0.en.html
* Original source code from Debao Zhang(2014) : https://github.com/dbzhang800/QtXlsxWriter
* Qt License (see Qt Company site) : https://www.qt.io/
   
## Contact
* [https://github.com/j2doll](https://github.com/j2doll)
* [j2doll@gmail.com](mailto:j2doll@gmail.com)
* Hi! My native language is not English. My English is not fluent. Please use EASY English. :-) 
