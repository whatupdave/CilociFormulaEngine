FormulaEngine
-------------

Version: [Version]
Release Date: [Date]
Author: Eugene Ciloci

What is it?
-----------
A .NET assembly that implements a formula engine.  It manages the parsing of formulas, tracking their dependencies, and recalculating in natural order.  The formula syntax is based on Excel and the library comes with many common Excel functions already implemented.

What license does it use?
-------------------------
The library is licensed under the LGPL.

What's in the release?
----------------------
The release contains the following:

lib - The actual assembly with xml and chm documentation
demo - A poor man's Excel implementation showing off the engine's features
src - Source code for the library, sample, and test framework

Dependencies
------------
The library uses the excellent grammatica parser generator to handle formula parsing.
It is also licensed under the LGPL and you can get the latest version at: http://grammatica.percederberg.net/
The demo uses the ZedGraph chart component (included in release)
The tests use NUnit and the Excel COM object.

Additional Information
----------------------
The project is hosted on SourceForge.  You can find the project page at: http://sourceforge.net/projects/formulaengine

There is also an article about it on CodeProject:
http://www.codeproject.com/vb/net/FormulaEngine.asp

I hope you find it useful
