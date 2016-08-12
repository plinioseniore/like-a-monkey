# Data Mining

## Introduction

The following set of Macro and Matlab/Octave function allow to mine data and provide a quick automated way to process the classification of data.

Naming and schema are designed to work for datalist for Industrial Automation projects.

## Excel Macro

### PropertyTypeCount Function

This function is sued to identify a Target property into a CellRange that applies for a LoopNumber

```
Public Function PropertyTypeCount(LoopNumber As String, CellRange As Range, Target As String) As Variant
```

|   A      |    B     |  C  |
|---|---|---|
|Tagname   |LoopNumber|Type |
|PT-941001 |P-941001  | AIR |
|PV-941001A|P-941001  | AOR |
|PV-941001B|P-941001  | AOR |

```LoopTypeCount("P-941001", B:C, "AO")``` **returns 2**. 

```LoopTypeCount("P-941001", B:C, "AI")``` **returns 1**.

### LoopNumber Function

This function returns the associated LoopNumber starting from a TagName.

```
Public Function LoopNumber(TagName As String, WithLoopType As Boolean) As String
```

```LoopNumber("PT-941001", FALSE)``` **returns 941001.**
```LoopNumber("PT-941001", TRUE)```  **returns P941001.**

### RetrieveTagNumber Function

This function retrieve a TagNumber (or TagName) from a general string, it assume that the TagNumber is delimited by "space" character before and after the TagNumber it-self.

```
Public Function RetrieveTagNumber(StringValue As String, UnitNumber As String) As String
```

Look for a tagnumber in the form of **LLLLLUUNNNN** where:
- **LLLL** is the tag type (FT, FI, FAL, FALL, ...)
- **UU**   is the unit number
- **NNNN** is the loop number

```RetrieveTagNumber("IN DIFFERENT CARD THAN FT941001 TRANSMITTER", "94")``` **returns FT941001.**

## Matlab/Octave Function

### KNNsearch Function

This function is used to *classify* a set of data using a **trained** set of examples. Returns the *classification* result and the *distance* from the target class.

```
function [classified, dist] = KNNsearch(trained, data, k, distance)
```

The KNN algorithm can process only numerical data, so a *model* of categorical data is required. An example of classification of categorical data is available in the [KNN Classification Example](knn_classification/README.md).
