# **xlFOCUS (v. 0.3)**																		
																		
This spreadsheet exposes examples of how to fetch data from the BCB's webservices, including FOCUS (market expectations), SGS (economic indicators), SCR (credit data), and SPI (payments system).																		
It is intended to be used by researchers and the general public. It is NOT a product of the BCB, nor it is maintained by that institution. Use at your own risk!																		
It is totally free and its code is open!																		
																		
## **TIPS**																		
																		
* Most recent version should be found in the webpage below:																		
https://github.com/edugca/xlFOCUS																		
																		
* To update from a previous version, just replace the old version of this spreadsheet with this one.																		
To update in your own spreadsheets, you need to open each one of them in the VBA Editor and replace the module "xlFOCUS" with its current version embedded in this spreadsheet.																		
																		
* To use these functions in your own spreadsheet, use this spreadsheet as a model, it is EASIER!																		
Alternatively, make a reference in your spreadsheet to the "Microsoft Scripting Runtime" library and copy to your spreadsheet the VBA modules named xlFOCUS and JsonConverter, both embedded in this spreadsheet.																		
The JsonConverter module is part of the "VBA-JSON" project developed by Tim Hall and available at the website below. Tested on v2.3.1.																		
https://github.com/VBA-tools/VBA-JSON																		
																		
	1) With your spreadsheet open, make sure macros are enabled and the Developer tab is displayed. In the links below, you learn how to do that:																	
																		
	https://support.microsoft.com/en-us/office/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6																	
	https://support.microsoft.com/en-us/topic/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45																	
																		
	2) In the ribbon, click on Developer tab, then on "Visual Basic".																	
																		
	3) In the toolbar of the VBA Editor, click on Tools, then on References. Enable the library "Microsoft Scripting Runtime". Then click OK.																	
																		
																		
	4) Drag the VBA modules "xlFOCUS" and "JsonConverter" from this spreadsheet to your spreadsheet. Now, close the VBA Editor. You're ready to use xlFOCUS!																	
																		
																		
* These functions rely on Excel's WEBSERVICE function (available from Excel 2013 onward), which can be quite unstable. So, when designing your spreadsheet, remember to save it constantly.																		
																		
* There should be a function for each resource provided by the BCB and covered by this tool. One can find the metadata on their webpages:																		
FOCUS	https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/aplicacao#!/recursos																	
SCR	https://olinda.bcb.gov.br/olinda/servico/taxaJuros/versao/v2/aplicacao#!/recursos																	
SGS	https://dadosabertos.bcb.gov.br/dataset/20542-saldo-da-carteira-de-credito-com-recursos-livres---total/resource/6e2b0c97-afab-4790-b8aa-b9542923cf88																	
SPI	https://olinda.bcb.gov.br/olinda/servico/SPI/versao/v1/aplicacao#!/recursos																	
																		
* If your query is not working, try to build it on the webpage above. If the server is down, then all functions will fail to fetch data!																		
																		
* Beware that the BCB's  webservices are case sensitive, that is, "IPCA" works but "ipca" does not.																		
																		
* To read large JSON files generated by the webservices (tested with FOCUS, SCR and SPI), download them to a folder and then read them with the function xlFOCUS_ReadJSONFile.																		
																		
* There are some known limitations with functions that directly query the webservices. Basically, avoid returning much data in a single function call and avoid many function calls:																		
	* 10,000 is the maximum number of observations returned by each function call																	
	* 32,767 is the maximum number of characters of the JSON script returned by each function call																	
	* 2,048 is the maximum number of characters of the URL path accessed by each function call																	
																		
## **Version history**																	
																		
v 0.1 (2021-12-16)

* First release.																		
																		
v 0.2 (2021-12-19)

* More instructions on how to implement these resources on the user's spreadsheet.												
* New function to read JSON files: xlFOCUS_readJSONFile.																		
																		
v 0.3 (2021-12-26)

* Fixed the encoding of the text read by the function xlFOCUS_ReadJSONFile. Now, it is correctly set to UTF-8.									
* New function to get data from the SGS system: xlFOCUS_SGS.															
* New function to read JSON script returned from the SGS system: xlFOCUS_SGS_ReadJSON.												
* New function to get data from the SCR system: xlFOCUS_SCR_TaxasDeJurosDiario.													
