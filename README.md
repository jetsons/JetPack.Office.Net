# JetPack.Office for .NET

To use this simply grab our Nuget package `Jetsons.JetPack.Office` and add this to the top of your class:

    using Jetsons.JetPack.Office;
	
This statement unlocks all the extension methods below. Enjoy!

This library depends on the following Nuget packages:

- Jetsons.JetPack
- TikaOnDotNet
- DocumentFormat.OpenXml
- EPPlus
	
### Extensions

Extension methods for file I/O performed using file path Strings:

- string.**LoadXLSX**
- string.**LoadRTFAsText**
- string.**LoadDOCAsText**
- string.**LoadDOCXAsText**
- string.**LoadDOCXAsTextFast**
- string.**LoadPDFAsText**
- string.**LoadXLSAsText**
- string.**LoadXLSXAsText**
- string.**LoadPPTAsText**
- string.**LoadPPTXAsText**