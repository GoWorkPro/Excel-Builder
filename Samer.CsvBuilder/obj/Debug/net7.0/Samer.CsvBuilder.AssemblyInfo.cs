//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.42000
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Reflection;

[assembly: System.Reflection.AssemblyCompanyAttribute("GoWorkPro")]
[assembly: System.Reflection.AssemblyConfigurationAttribute("Debug")]
[assembly: System.Reflection.AssemblyCopyrightAttribute("CsvBuilder Utility © 2023 Samer Shahbaz. All rights reserved.")]
[assembly: System.Reflection.AssemblyDescriptionAttribute("\t\tVery simple and Easy to use, convert datatables to CSV\r\n\t\tAuthor: Samer Shahbaz" +
    "\r\n\r\n\t\tCreate Date: 15/11/2023\r\n\r\n\t\tDescription:\r\n\t\tThe CsvBuilder utility, devel" +
    "oped by Samer Shahbaz, is a powerful tool designed to simplify the process of cr" +
    "eating CSV (Comma-Separated Values) files using .NET DataTables. This utility pr" +
    "ovides a convenient and efficient way to generate CSV data from one or more Data" +
    "Tables within a DataSet.\r\n\r\n\t\tKey Features:\r\n\r\n\t\tDataSet Integration: Accepts a " +
    "DataSet as input, allowing the user to aggregate multiple DataTables for CSV cre" +
    "ation.\r\n\t\tFlexible Value Rendering: Supports a customizable event, ValueRenderEv" +
    "ent, which allows users to define a custom parser for values based on their data" +
    " type (column or row).\r\n\t\tMultiple DataTable Support: Enables the user to select" +
    "ively include columns from different DataTables by specifying the table index.\r\n" +
    "\t\tStream Handling: The utility efficiently manages memory streams to optimize CS" +
    "V generation.\r\n\t\tDispose Method: Implements the IDisposable interface for proper" +
    " resource management.\r\n\t\tUsage:\r\n\r\n\t\tConstructor:\r\n\t\tStatic Method for Creating " +
    "CsvBuilder with Multiple DataTables:\r\n\t\tCsvBuilder csvBuilder = CsvBuilder.Datas" +
    "ets(dataTable1, dataTable2, ...);\r\n\t\r\n\t                Building CSV:\r\n\t\tcsvBuild" +
    "er.Build(tableIndex1, tableIndex2, ...)\r\n\t\tCustomizing Value Rendering:\r\n\r\n\t\tSub" +
    "scribe to the ValueRenderEvent to define custom parsing logic for column and row" +
    " values.\r\n\t\t\r\n                                Output Handling:\r\n\t\tObtain the CSV" +
    " content as a Stream:\r\n\t\tStream csvStream = csvBuilder.GetStream();\r\n\t\tSave the " +
    "CSV content to a file:\r\n\r\n\t\tcsvBuilder.SaveAsFile(\"filePath.csv\");\r\n\t\t\r\n        " +
    "                         Example #1:\r\n\r\n\t\t// Create CsvBuilder with a DataSet\r\n\t" +
    "\tICsvBuilder csvBuilder = CsvBuilder.Datasets(dataSet);\r\n\r\n\t\t// Build CSV with s" +
    "elected columns from specific DataTables\r\n\t\tICsvExtractor csvExtractor = csvBuil" +
    "der.Build();\r\n\r\n\t\t// Obtain CSV content as a Stream\r\n\t\tMemoryStream csvStream = " +
    "csvExtractor.GetStream();\r\n\r\n\t\t// Save CSV content to a file\r\n\t\tcsvExtractor.Sav" +
    "eAsFile(\"output.csv\");\r\n\r\n\t\t//dispose if necessary\r\n\t\tcsvExtractor.Dispose();\r\n\t" +
    "\tThis utility simplifies the process of CSV generation, providing users with a f" +
    "lexible and efficient solution for working with tabular data in the .NET environ" +
    "ment.")]
[assembly: System.Reflection.AssemblyFileVersionAttribute("2.0.0.0")]
[assembly: System.Reflection.AssemblyInformationalVersionAttribute("2.0.0")]
[assembly: System.Reflection.AssemblyProductAttribute("Samer.CsvBuilder")]
[assembly: System.Reflection.AssemblyTitleAttribute("Samer.CsvBuilder")]
[assembly: System.Reflection.AssemblyVersionAttribute("2.0.0.0")]

// Generated by the MSBuild WriteCodeFragment class.

