# VBA.Round
Functions for rounding Currency, Decimal, and Double up, down, by 4/5, or to a specified count of significant figures.

In many areas, rounding that accurately follows specific rules are needed - accounting, statistics, insurance, etc.

Unfortunately, the native functions of VBA that can perform rounding are either missing, limited, inaccurate, or buggy, and all address only a single rounding method. The upside is that they are fast, and that may in some situations be important.

However, often precision is mandatory, and with the speed of computers today, a little slower processing will hardly be noticed, indeed not for processing of single values. All the functions presented here run at about 1 µs.

They cover the normal rounding methods:

•Round down, with the option to round negative values towards zero

•Round up, with the option to round negative values away from zero

•Round by 4/5, either away from zero or to even  (Banker's Rounding)

•Round to a count of significant figures

The first three functions accept all the numeric data types, while the last exists in three varieties - for Currency, Decimal, and Double respectively.
They all accept a specified count of decimals - including a negative count which will round to tens, hundreds, etc. Those with Variant as return type will return Null for incomprehensible input.
   
Files can be imported into an existing project with the command:

LoadFromText acModule, "RoundingMethods", "d:\path\RoundingMethods.bas"


Documentation is in-line or here:

CodePlex:
http://www.codeproject.com/Tips/1022704/Rounding-Values-Up-Down-By-Or-To-Significant-Figur

Experts-Exchange:
http://rdsrc.us/K5cO9F

