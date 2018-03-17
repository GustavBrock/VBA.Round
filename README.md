# VBA.Round

### Introduction

Superior functions for:

* Rounding Currency, Decimal, and Double up, down, by 4/5, or to a specified count of significant figures
* Rounding \(and scaling\) all items of a series of numbers to have the sum to match a desired rounded total

In many areas, rounding that accurately follows specific rules are needed - accounting, statistics, insurance, etc.

Unfortunately, the native functions of VBA that can perform rounding are either missing, limited, inaccurate, or buggy, and all address only a single rounding method. The upside is that they are fast, and that may in some situations be important.

However, often precision is mandatory, and with the speed of computers today, a little slower processing will hardly be noticed, indeed not for processing of single values. All the functions presented here run at about 1 µs.

They cover the normal rounding methods:

* Round down, with the option to round negative values towards zero
* Round up, with the option to round negative values away from zero
* Round by 4/5, either away from zero or to even  \(Banker's Rounding\)
* Round to a count of significant figures

The first three functions accept all the numeric data types.  
The forth exists in three varieties - for Currency, Decimal, and Double respectively.

Finally, the function **RoundSum** will round a series of numbers so the sum of these matches the rounded sum of the unrounded values. Further, if a requested total is passed, the rounded values will be _scaled_, so the sum of these matches the rounded total. In cases where the sum of the rounded values doesn't match the rounded total, the rounded values will be adjusted where the applied error will be the relatively smallest.

They all accept a specified count of decimals - including a negative count which will round to tens, hundreds, etc. Those with Variant as return type will return \_Null\_ for incomprehensible input.

### Usage

Files can be imported into an existing VBA project with the command:

`LoadFromText acModule, "RoundingMethods", "d:\path\RoundingMethods.bas"`

### Documentation

Documentation is in-line. Articles on the topic can be found here:

**CodePlex**:  
[http://www.codeproject.com/Tips/1022704/Rounding-Values-Up-Down-By-Or-To-Significant-Figur](http://www.codeproject.com/Tips/1022704/Rounding-Values-Up-Down-By-Or-To-Significant-Figur)

**Experts-Exchange**:

[Rounding values up, down, by 4/5, or to significant figures](https://www.experts-exchange.com/articles/20299/Rounding-values-up-down-by-4-5-or-to-significant-figures.html)
[Round elements of a sum to match a total](https://www.experts-exchange.com/articles/31683/Round-elements-of-a-sum-to-match-a-total.html)

