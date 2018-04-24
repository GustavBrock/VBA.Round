# VBA.Round
---

### Introduction

Superior functions for:

* Rounding Currency, Decimal, and Double up, down, by 4/5, or to a specified count of significant figures
* Rounding \(and scaling\) all items of a series of numbers to have the sum to match a desired rounded total

Also:

* Rounding by the Power of Two (Base 2), up, down, or by 4/5
* Convert and format imperial distance (feet and inches) with high precision


## General rounding
![General Rounding](/images/EE Round.png)

In many areas, rounding that accurately follows specific rules are needed - accounting, statistics, insurance, etc.

Unfortunately, the native functions of VBA that can perform rounding are either missing, limited, inaccurate, or buggy, and all address only a single rounding method. The upside is that they are fast, and that may in some situations be important.

However, often precision is mandatory, and with the speed of computers today, a little slower processing will hardly be noticed, indeed not for processing of single values. All the functions presented here run at about 1 µs.

They cover the normal rounding methods:

* Round down, with the option to round negative values towards zero. Base 10 and Base 2
* Round up, with the option to round negative values away from zero. Base 10 and Base 2
* Round by 4/5, away from zero. Base 10 and Base 2
* Round by 4/5, to even  \(Banker's Rounding\). Base 10
* Round to a count of significant figures. Base 10

The first three functions accept all the numeric data types.  
The forth exists in three varieties - for Currency, Decimal, and Double respectively.

All functions accept a specified count of decimals - including a negative count which will round to tens, hundreds, etc. Those with Variant as return type will return *Null* for incomprehensible input.

## Rounding a series of numbers to a sum

![Rounding a series of numbers to a sum](/images/EE Slices.png)

The function **RoundSum** will round a series of numbers, so the sum of these matches the rounded sum of the unrounded values. Further, if a requested total is passed, the rounded values will be _scaled_, so the sum of these matches the rounded total. In cases where the sum of the rounded values doesn't match the rounded total, the rounded values will be adjusted where the applied error will be the relatively smallest.

## Rounding by a power of two

![Rounding by a power of two](/images/EE Power 2.png)

This will not round by:

	1000, 100, 10, 1, 1/10, 1/100, 1/1000, etc. 

as for decimals, but by a power of two:

	32, 16, 8, 4, 2, 1, 1/2, 1/4, 1/8, 1/16, 1/32, etc.

again, with extreme precision (down to 2<sup>-21</sup>) and including very small or very large numbers (2<sup>-96</sup> to 2<sup>96</sup>).

## Converting between meters and inches

![Converting between meters and inches](/images/EE Imperial.png)

A practical usage of *rounding by the power of two* is to convert back and forth between metric and imperial measures:

    meters <=> feet, inches, fractions
    
The functions provided will handle very large and very small values for inches:

    from ±7922816299999618530273437599 
    to 1/2097152 or the decimal value 0.000000476837158203125

The format of the imperial output covers a very wide range:

* Feet and inches, or inches only
* No fraction for a numerator of zero
* No fraction at all
* Dash or no dash between feet and inches
* Only feet if total of inches is 12 or more
* Zero feet if total of inches is smaller than 12
* No inches if feet are displayed and inches are zero
* No units
* Units spelled out as ft, ft., or foot/feet and in, in., or inch/inches

and many variations hereof.

## Usage

Files can be imported into an existing VBA project, for example for module *RoundingMethods* with the command:

    LoadFromText acModule, "RoundingMethods", "d:\path\RoundingMethods.bas"

Likewise for the other modules.  

## Documentation

Detailed documentation is in-line. 

Articles on the topic can be found here:
 
![CP Logo](/images/CP Logo Small.png)

[Rounding values up, down, by 4/5, or to significant figures](http://www.codeproject.com/Tips/1022704/Rounding-Values-Up-Down-By-Or-To-Significant-Figur)


![EE Logo](/images/EE Logo.png)
 
[Rounding values up, down, by 4/5, or to significant figures](https://www.experts-exchange.com/articles/20299/Rounding-values-up-down-by-4-5-or-to-significant-figures.html)

[Round elements of a sum to match a total](https://www.experts-exchange.com/articles/31683/Round-elements-of-a-sum-to-match-a-total.html)

[Round by the power of two](https://www.experts-exchange.com/articles/31859/Round-by-the-power-of-two.html)

[Convert and format imperial distance (feet and inches) with high precision](https://www.experts-exchange.com/articles/31931/Convert-and-format-imperial-distance-feet-and-inches-with-high-precision.html)
