# xlFOCUS
xlFOCUS: fetch BCB's FOCUS survey data to Excel

This spreadsheet exposes examples of how to fetch data from the BCB's "Expectativas de Mercado" survey, also known as FOCUS.
It is intended to be used by researchers and the general public. It is NOT a product of the BCB, nor it is maintained by that institution. Use at your own risk!
It is totally free and its code is open!

TIPS															
															
* To use these functions in your own spreadsheet, copy to the latter the VBA modules named xlFOCUS and JsonConverter, both embedded in this spreadsheet.
The JsonConverter module is part of the "VBA-JSON" project developed by Tim Hall and available at the website below. Tested on v2.3.1.

https://github.com/VBA-tools/VBA-JSON								
															
* These functions rely on Excel's WEBSERVICE function (available from Excel 2013 onwards), which can be quite unstable. So, when designing your spreadsheet, remember to save it constantly.															
															
* There is a function for each resource provided by the FOCUS's webservice. One can find the metadata on its webpage:
	
https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/aplicacao#!/recursos																		
* If your query is not working, try to build it on the webpage above. If the server is down, then all functions will fail to fetch data!														
* Beware that the FOCUS webservice is case sensitive, that is, "IPCA" works but "ipca" does not.

* There are some known limitations. Basically, avoid returning much data in a single function call:															
	* 9,999 is the maximum number of observations returned by each function call														
	* 32,767 is the maximum number of characters of the JSON script returned by each function call														
	* 2,048 is the maximum number of characters of the URL path accessed by each function call
