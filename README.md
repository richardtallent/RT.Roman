Update 2019-03-25
=================
Apparently I failed to notice that Microsoft recently added the ARABIC() and ROMAN() functions to do exactly what this library does. I'm keeping this here in case anyone needs to support older versions of Excel, but since I'm using fairly modern Excel versions at work and home, I won't be making any changes to this library.

Purpose
=======
Latin may be dead for most humans, but Roman numeral are still used for everything from document outlines to generational names to seismology. My own needs stemmed from working with legal documents in Excel. I couldn't find a decent set of functions to convert back and forth between decimal and Roman, so I rolled my own in VBA (another language that's not quite dead yet).

Status
======
This library, to my knowledge, is stable, accurate, and as performant as is practical in VBA. Suggestions welcome.

I'm considering a port to C#, since .NET also lacks native Roman/decimal conversion. I already have the code ported, but not extracted to its own library, and the algorithm is a bit different.

Revision History
================
 - 1.0		2017-04-02	First release

Using this Library
====================
Import the `modRoman.bas` module file into your Excel workbook. This provides two functions, which you can either call from VBA or as user-defined functions directly from your cell formulas:

### RomanToDecimal
Pass it a Roman numeral string, it will return the integer equivalent.

```Visual Basic
RomanToDecimal("MMMDCLVIII")	' Returns 3658
```

### DecimalToRoman
Pass it an integer, it returns an uppercase Roman numeral equivalent. (You can lowercase it yourself if you wish, native lowercase support isn't built in yet.)

```Visual Basic
DecimalToRoman(2017)	' Returns MMXVII
```

Large Number Limitations
========================
I limited support to positive 16-bit integer values, i.e. 1 - 16,384. My use case would rarely have numbers even exceeding 100, so this seemed a reasonable limit.

For academic reasons only, I considered adding support for ↁ (5,000) and ↂ (10,000), since these Unicode endpoints have strong font support, but I haven't done so yet.

It may also be possible to add *viniculum* support using the Unicode combining overline character 0x305 (so, for example, 25,000 could be represented as X&#x305;X&#x305;V&#x305;), but this is not yet supported, and neither the overline nor macron combining characters produce a full-width line.

Other Caveats
=============
 - Zero is not supported, nor was it in ancient Rome. I'd be open to supporting "nulla" or "N" (which came around in the 6th century CE or so) if someone feels strongly about it.
 - Negative numbers are not supported because the Romans had no concept for negative numbers. Adding an Arabic negative sign would be easy, again, if someone really wants it.
 - Unicode endpoints specifically designed to represent Roman numerals are not used or converted, only Latin letters I, V, X, L, C, D, and M.
 - Invalid strings sent to `RomanToDecimal` return 0. The code could be modified to raise an error instead.

Performance
===========
I've seen a number of other conversion functions, both in VBA and other languages, and far too many use relatively slow mechanisms such as regular expressions and string substitution.

Performance was important to me, so I've optimized the code as best I can without resorting to esoterics (such as using array manipulation over string concatenation).

If you see an opportunity to further optimize without detracting from code readability, please feel free to issue a pull request! All code submitted should be compatible with Excel 2007+ on Windows and Excel 2011+ on OS X.

`RomanToDecimal` works by parsing the string right to left, one character at a time. This significantly reduces the complexity of determining whether an I, X, or C are *adding* or *subtracting* from the value. If it finds that the string contains an invalid sequence (e.g., "DD"), it will fail and return 0.

`DeimalToRoman` works from the thousands place(s) down to the ones. The hundreds, tens, and ones have exceptions for their 4 and 9 values, so those are handled first, and those values can skip the tests for the 5 and 1 values. As a side note, I tested using `String$()` to compute the multiples of `C`, `X`, and `I` (as I did with `M`), but it was 20% slower than the `Select Case` statements.

License (MIT "Expat")
=====================
Copyright 2017 Richard S. Tallent, II

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
