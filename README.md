# Quandl-Grapher README #

# Introduction #
Quandl Grapher is an Excel Add-in that allows Excel and VBA users to query financial data from the [https://www.quandl.com/](https://www.quandl.com/ "Quandl API") and present the data in graphs and other formats. The software acts as a wrapper for [https://github.com/HubertJ/QuandlCS](https://github.com/HubertJ/QuandlCS "QuandlCS") along with the graphing capabilities found in [Excel's Chart Interface ](https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.chart.aspx "Excel's Chart Interface"). 

The Quandl-Grapher dll is especially powerful when interfaced with VBA.

The goal of this Excel Add-in is to reproduce graphing functionality typically found on websites and Bloombergs. By having this on Excel, everyday investors can manipulate and organize their data using Excel.

# Documentation #
Libraries that enabled this: <br/>
- [http://exceldna.codeplex.com/](http://exceldna.codeplex.com/ "Excel-DNA") <br/>
- [https://github.com/HubertJ/QuandlCS](https://github.com/HubertJ/QuandlCS "QuandlCS")<br/>

# Usage #
## Excel Spreadsheet ##

## VBA Usage ##
One of the possible usage cases is interfacing with VBA.

    Sub graphing()
	    Dim lib As Object: Set lib = CreateObject("QuandlGrapher")
	    Dim o As Variant
	    o = lib.DataSet("YAHOO", "INDEX_GSPC", 100, "d")
	    o = lib.CreateSingleGraph()
	End Sub

# Functionality #
## Graphing (QuandlGraphing) ##
The following output is the simplest graph that can be produced: <br/>
![](https://cloud.githubusercontent.com/assets/6467461/10484955/1f45f0c8-7256-11e5-850c-cc6eee3862c3.jpg) <br/>
More functionality will be coming, with rolling averages, and bollinger bands.

I hope to implement code that will identify trends such as shoulders and heads, double/triple tops and bottoms, and other technical charting methods.

## Outputting Raw Data (DataAccess) ##

## Options (QuandlOptions) ##

# Future Plans #