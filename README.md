# VBA-to-JavaScript-Translator
This translator is intended to be used as an educational tool to help VBA developers get familiar with JavaScript. 

The code for this tool is based on regex examples found from multiple sources in forums and online.  Sorry to say...I don't have the information needed to directly credit all the sources I leverageed.
However a thank you goes to the folks at Slingfive.com for the inspiration

As with most translation tools, this tool will NOT perform a 100% complete translation. 
It is designed to cover the most used constructs in VBA (enough to get you started). 

Start by entering or pasting a basic block of VBA code (be sure to include your Function/Sub wrappers). 
Review how that code looks in JavaScript. Next, try increasingly more advanced code examples (loops, if statements, etc.). 

Currently, this translator works with:
* Both Functions and Sub Procedures
* Variable Declarations
* Basic If Then Statements
* Select Case Statements
* Most Comparison Operators
* For Each Loops
* Most variations of Do Loops
* Basic MsgBox calls

This tool will NOT accurately translate:
*  Built in VBA functions
*  References to Excel Objects
*  For Each loops that use STEP constructs

I offer the source code up to anyone interested in helping makes this tool a more robust utility for the VBA community.

