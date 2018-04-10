html2xlsxefs
============

Based on html2xlsx with the following changes:

* tables can be given a title attribute to set the tab name
* text is wrapped
* money type added
* cell height is no longer set (needed for wrapped text)
* cell format can be set using the data-num-format attribute e.g. 'mmm yyyy'
