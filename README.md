# VBA-to-JavaScript-Translator
This translator is intended to be used as an educational tool to help VBA developers get familiar with JavaScript. 

The code for this tool is based on regex examples found from multiple sources in forums and online.  Sorry to say...I did not keep a record of all the sources I leveraged.  Just know that people smarter than I am provided the basis for most of the code found here.  I simply adjusted what I needed and slapped it all together into a single utility.

As with most translation tools, this tool will NOT perform a 100% complete translation. 
It is designed to cover common constructs in VBA (enough to get you started). 

The idea is to enter a basic block of VBA code (be sure to include your Function/Sub wrappers) and see how the syntax would look in JavaScript.  It's best if you first try something simple then progress into more advanced IF statements, loops, comparison operators, Select Case switches, etc. 

Currently, this translator works with:
* Both Functions and Sub Procedures
* Variable Declarations
* Basic IF THEN Statements
* SELECT CASE Statements
* Most Comparison Operators
* Basic FOR x TO y STEP z Loops
* Most variations of DO LOOPS
* Basic MSGBOX calls


Known Issues:
This tool currently does NOT accurately translate
*  Built in VBA functions
*  References to Office Objects
*  WITH Statements
*  FOR EACH Loops (in fact these cause the tool to return nothing at all.  I'm still working out why that is) 

I offer the source code up to anyone interested in helping make enhancements to this tool in order to develop a more robust utility for the VBA community.

See working tool here:  http://www.datapigtechnologies.com/VBAToJS/VBAToJavaScriptTranslator.html
