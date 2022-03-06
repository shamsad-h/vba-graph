# README

This is a module for Excel written in VBA designed to generate graphs for share price data.

The module works by taking two inputs from different cells in the spreadsheet:

* Data range (written in the format, e.g., A2:B10)
* Title for the graph

Key features:

* The graph automatically scales the axes:
- The y-axis automatically changes its maximum and minimum values to make the graph as large as possible (to aid with comparison with other graphs)
- The x-axis automatically changes between years and months depending on the timeframe
* The graph can accomodate share price data for as many stocks/indices as necessary
* The graph is set to format to a ready-to-use design