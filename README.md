# Barcode Encoding
Originally developed to add Code 128 barcodes to SSRS reports on a server that can't have anything custom installed on it.

The text from `Code128.vb` can be dropped in to the `<code>` section of an SSRS report.

Then textboxes where barcodes are desired can be ran through `GetCode128EncodedString` to get the text you would then display with a [Code 128 barcode font](http://www.barcodelink.net/code128.ttf).