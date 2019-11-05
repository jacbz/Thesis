# Thesis – An Excel to code converter

**Thesis** reads Excel files and builds a graph-based model to generate structured code from the procedures contained in each cell's formula. Right now, generating code in **C#** is supported. This project is written in C# / WPF.

![]( https://i.imgur.com/J3u4LMC.png )

Features include:

* "Intelligent" **variable name generation** from labels positioned inside the spreadsheet
* Support for **empty cells**
* Partial support for expressions that are not **type-safe**
* **Ranges** (such as A1:C3)
* **Names**
* References to other **worksheets**
* **Functions**
  * **Arithmetic**: SUM, MIN, MAX, COUNT, AVERAGE, ROUND, ROUNDUP, ROUNDDOWN
  * **Logical**: IF, NOT, AND, OR, XOR
  * **Reference**: VLOOKUP, HLOOKUP, CHOOSE, MATCH, INDEX
  * The code is structured in a way that makes implementation of other functions easy.

## Documentation

This application was created as part of a Bachelor's thesis at the Technical University of Munich. **It is not in active development anymore.**

The thesis is included in the file **[Thesis.pdf](Thesis.pdf)** and includes the ideas behind this project and some implementation details.

## Getting Started

### Prerequisites

I recommend using Visual Studio and NuGet.

* .NET Framework >= 4.7.2
* Windows 7 or higher (for WPF)

### Syncfusion license

This project uses proprietary components by [Syncfusion](https://www.syncfusion.com/):

* [SfSpreadsheet](https://help.syncfusion.com/wpf/spreadsheet/overview#key-features) is used to load, display and retrieve data from Excel spreadsheets.
* [SfDiagram](https://help.syncfusion.com/wpf/sfdiagram/overview) is used to draw the spreadsheet as a graph.

To use those components for free, you can potentially obtain a **[Community license](https://www.syncfusion.com/products/communitylicense)**, which as of now is eligible for companies and individuals with less than $1 million USD in annual gross revenue and 5 or fewer developers.

Once you have obtained a license key, put it inside the `Resources/SyncfusionLicenseKey.txt` file.

### Installing

1. Clone the project.
2. [Load](https://docs.microsoft.com/nuget/consume-packages/reinstalling-and-updating-packages) the required packages via NuGet.
3. Put your license inside the`Resources/SyncfusionLicenseKey.txt` file. **Do not commit this file!**
4. Ready!

## License

This project is licensed under the **[GNU General Public License, Version 3](https://www.gnu.org/licenses/gpl-3.0.de.html)**. See the LICENSE file for more details.

Components used include:

* [AvalonEdit](http://avalonedit.net/), a text box with syntax highlighting (MIT License)
* [Costura.Fody](https://github.com/Fody/Costura), to include all libraries into the output .exe (MIT License)
* [FontAwesome](https://github.com/MartinTopfstedt/FontAwesome5), for the button icons (CC BY 4.0 / MIT License)
* [MahApps.Metro](https://github.com/MahApps/MahApps.Metro), a WPF UI toolkit (MIT License)
* [Roslyn](https://github.com/dotnet/roslyn), for automated testing (Apache License 2.0)
* [XLParser](https://github.com/spreadsheetlab/XLParser), an Excel formula parser (Mozilla Public License 2.0)

The icon was made by [Freepik](https://www.freepik.com/home) from [www.flaticon.com](http://www.flaticon.com/).

## Automated testing

There is a second branch "ObjectOrientedApproach", which is an implementation of the Object-Oriented Approach as described in the thesis. It has since been superseded by the Function-based Approach, which is now the Master branch.

The Object-Oriented Approach contains an **automated tester**, which tests the generated code against the actual values as evaluated by the spreadsheet's formula engine (using the [Roslyn Scripting API](https://github.com/dotnet/roslyn/wiki/Scripting-API-Samples)). This has not been implemented for the Master branch yet.