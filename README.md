# PiBridgeForExcel

This is a demo excel workbook that fires order to Pi using PiBridge.
You first need to register the PiBridge.dll in System32/SysWOW64 folder, so that Excel can find this dll.

Download the .bat file and all other files from https://github.com/howutrade/PiBridgeForExcel/tree/master/PiBridge.dll and follow the notes to register the PiBridge.dll

In the demo excel workbook, we wrapped the PiBridge PlaceOrder function in a UDF with inbuilt validation.
