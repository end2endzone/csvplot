[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Github Releases](https://img.shields.io/github/release/end2endzone/csvplot.svg)](https://github.com/end2endzone/csvplot/releases)



# csvplot

csvplot is a windows command line script which uses Microsoft Excel to plot a graph from a comma separated values (*.csv) file

It's main features are:
*  No external libraries or compilation required.
*  Support multiple command line arguments.
*  Build on top of Microsoft Excel.
*  Supports lossless PNG image format.
*  Supports CSV data file format which is a generic text file format supported by any application and programming language.
*  Supports virtually an unlimited number of time series plots.
*  Automatically detect appropriate boundaries for the graph.

# Purpose

Most programming language does not offer a native library for plotting data to a graph. They often require external libraries for implementing the process and each library does not work the same way.

Excel is a native platform for parsing Comma-separated values (CSV) files and can be scripted to plot the result into an image.

The purpose of this script is to allows any programming language which generates data to plot the data into an image by saving the raw data to a CSV file and then using the power of Excel to plot the result into an image.


# Limitations:

The script has some limitations which are explained here.

## Output image

The resolution of the output image may be +- 1 pixel different than what is requested on command line. This is a limitation of how Excel processes image dimensions since it uses "points" as base unit and not actual pixels. A conversion from pixels to points must be calculated which may contains small accuracy error.

The only supported image format is PNG. It is still unknown if Excel actually support JPG for exporting graphs but PNG seems to be the perfect candidate since its a lossless compressed format.

## Column Titles

It is expected that first row of each column contains the title of the column which will be used as the name of the plotted series within the graph.


# Usage

## Execute (command line)

The script is written in [VBScript](http://en.wikipedia.org/wiki/VBScript) and must be launched with the VBScript interpreter:
* `cscript` (for command line)
* `wscript` (for window mode)

The command for launching the script is as follows:

```batch
cscript //nologo csvplot.vbs [InputFile] [OutputFile] [Width] [Height] x1 y1 x2 y2 x3 y3
```

for example:

```batch
@echo off
cscript //nologo "%~dp0csvplot.vbs" path\to\demo.csv path\to\demo.png 800 600 1 2
pause
```

The script must be called with a minimum of 6 command line arguments:

| Index | Name                     | Description                                  |
|:-----:|--------------------------|----------------------------------------------|
|  1    |  InputFile               | Path of the input CSV file                   |
|  2    |  OutputFile              | Path of the output PNG image                 |
|  3    |  Width                   | Width of the output image in pixels          |
|  4    |  Height                  | Height of the output image in pixels         |
|  5    |  Serie #1, X column      | X column index of first series               |
|  6    |  Serie #1, Y column      | Y column index of first series               |
|  7    |  Serie #2, X column      | X column index of second series              |
|  8    |  Serie #2, Y column      | Y column index of second series              |
|  9    |  Serie #n, X column      | ...                                          |
| 10    |  Serie #n, Y column      | ...                                          |


**Note that column indice are 1-based and not 0-based. This means that column A is column 1 and not column 0.**

## Output

On execution, the following output is produced by the script:

```
Microsoft Windows [Version 6.1.7601]
Copyright (c) 2009 Microsoft Corporation.  All rights reserved.
 
C:\>cd /d C:\Temp\csvplotdemo
 
C:\Temp\csvplotdemo>cscript //Nologo csvplot.vbs %cd%\CarEngineModel.csv %cd%\CarEngineModel.png 853 479 1 3 1 4 1 2
Loading input file C:\Temp\csvplotdemo\CarEngineModel.csv...
File load successful.
File has 4 columns.
Plotting series of columns 1 and 3...
Plotting series of columns 1 and 4...
Plotting series of columns 1 and 2...
File C:\Temp\csvplotdemo\CarEngineModel.png saved successfully.
 
C:\Temp\csvplotdemo>
```

## Samples

The following section shows some example of using cvsplot to plot a series.

### Apple Share Prices

The following example show the Apple Share Prices closing value over the year 2015. The data is provided by Nasdaq at the following address: [http://www.nasdaq.com/symbol/aapl/historical](http://www.nasdaq.com/symbol/aapl/historical).

The CSV data can be downloaded here: [Apple Share Prices over time (2015).csv](/tests/Apple%20Inc.%20Common%20Stock%20Historical%20Stock%20Prices/Apple%20Share%20Prices%20over%20time%20(2015).csv)

[![Apple Share Prices over time](/docs/Apple%20Share%20Prices%20over%20time%20(2015).png)](/docs/Apple%20Share%20Prices%20over%20time%20(2015).png)

### Car Engine Model

The following example show a hypothetical car engine model which speed increase of decrease over time based on the feedback of the gas pedal.

The CSV data can be downloaded here: [CarEngineModel.csv](/tests/SinXCosXLogX/sinxcosxlogx.csv)

[![Car Engine Model](/docs/CarEngineModel.png)](/docs/CarEngineModel.png)

### Sin(x), Cos(x) and Log(x)

The following example show a graph of sin(), cos() and log() function in Excel.

The CSV data can be downloaded here: [sinxcosxlogx.csv](/tests/CarEngineModel/CarEngineModel.csv)

[![SinXCosXLogX](/docs/sinxcosxlogx.png)](/docs/sinxcosxlogx.png)

# Build / Install

N/A

# Testing

N/A

# Compatible with

csvplot is only available for the Windows platform and has been tested with the following version of Windows:

*   Windows 7

# Versioning

We use [Semantic Versioning 2.0.0](http://semver.org/) for versioning. For the versions available, see the [tags on this repository](https://github.com/end2endzone/csvplot/tags).

# Authors

* **Antoine Beauchamp** - *Initial work* - [end2endzone](https://github.com/end2endzone)

See also the list of [contributors](https://github.com/end2endzone/csvplot/blob/master/AUTHORS) who participated in this project.

# License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
