String Sort Benchmarks (26/12/2006)
by Ralph Eastwood (C) 2006

LICENSE

       This software is provided 'as-is', without any express or implied warranty.
       In no event will the authors be held liable for any damages arising from the use of this software.
       Permission is granted to anyone to use this software for any purpose, including commercial applications,
       and to alter it and redistribute it freely, subject to the following restrictions:

           - do not claim that this file, in source-form or binary-form is solely your work.
           - give acknowledgement for the use of the file if it forms an integral part of your code.
           - include these comments in any source-distribution you release.
                                                                            

INTRODUCTION

I have implemented two of the algorithms 'popular' MSD Radix Sorts 
based on the implementations done by Peter M. McIlroy and Keith Bostic in
their paper Engineering Radix Sort. Of the techniques used were VBSpeed techniques
such as safearray 'hacking' and others.
These algorithms go extremely fast, about 10x faster than a standard iterative version of
a QuickSort algorithm. This is shown the the accompanying benchmark...

Update: 03/01/2007
QuickSort algorithm replaced with TriQuickSort algorithm based on Phillipe Lord's version.
It goes faster than QuickSort, but still not as fast as the Radix Sorts. Thanks Steppenwolfe 
and RB about my faulty version.

Update: 04/01/2007
Small bug in TriQuickSort algorithm fixed.
Optimised TriQuickSort to be 2x faster.
Slight Optimisations.
Included a array sorted check routine.

Update: 04/02/2007
Non-recursive version of TriQuickSort
Included rd's stable non-recursive quick sort algorithm
Disabled assume no aliasing option in compile -> will cause errors.
Add Benchmark Exports To CSV
Can cancel during a long sort.

Update: 29/03/2007
CSV Export Modified (Little bug fix and different layout for the Simple Benchmark)
Included Ulli's Radix Sort and Quickie Sort

Update: 30/03/2007
Ulli's Radix Fixed, it was used incorrectly.
Allowed comparison of selected sorts

TUTORIAL

The interface should hopefully be very straight forward.

Simple Benchmark:
Count 		- The number of elements in an array to sort
String Length 	- The length of each string to be sorted
Deviate by 		- Have a random modifier for each string length when they are generated

Line Benchmark:
Initial Count	- The number of elements in an array to sort for first iteration
Final Count		- The number of elements in an array to sort for last iteration
Increments		- How many elements to increase the array after each sort
String Length 	- The length of each string to be sorted
Deviate by 		- Have a random modifier for each string length when they are generated

File Benchmark:
Click benchmark and it will prompt for a dictionary file.
A dictionary file can be any text file with words delimited with a vbCrLf a.k.a. newline
The benchmark will load the text file then randomise the word locations.
It will run benchmarks with each of the sorts and output their sorted results in the Output Folder.
This benchmark will take some time for large files, but progress is shown in the titlebar.

TODO

Add case-insensitive sorts.

CREDITS

Engineering Radix Sort, by Peter M. McIlroy and Keith Bostic\
Rde for his strSwapTable
Ulli for his Radix Sort and Quickie Sort Implementation

Ralph Eastwood