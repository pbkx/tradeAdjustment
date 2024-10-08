# tradeAdjustment

#### IMPORTS AND OTHER REQUIREMENTS

To install Pyap, run the command:
> pip install pyap

The program requires a file to pull addresses from, and a file that stores naming conventions. These are called "taa_2015_2017.xls" and "conventions.xls," respectively.

#### FUNCTIONALITY 

This method works using a parser to split an address into parts (think street number, street name, city, state, etc). 

Then, it individually standardizes each part using Python string processing algorithms.

First, it Pyap's parser to split an address into parts. Then, it deals with each of those parts individually. 

-> Some addresses contain numbers in word form, such as "One" instead of 1. These need to be converted back to numerical form.

-> Many addresses have strings such as "Street" that need to be converted to "ST," or "Highway" to "HWY." 

-> Some addresses contain a suite number or floor number, which needs to be appended to the end of the address.

If the parser cannot parse through an address (for example, an address with two numbers in it, such as "54 Highway 37," the program will pass that address to the "diff" function. This function deals with complicated addresses.

-> Some addresses are solely a postage box. 

-> Some addresses take the form of "## Road Name ##."

If all of this fails, the program returns "UNFORMATTED."

The standardized addresses are returned to an excel file called "formatted.xls."
