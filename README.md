# SSRS Barcode Encoding
Originally developed to add Code 128 barcodes to SSRS reports on an Epicor cloud server that is out of customer control.

## Only compatible with a specific font
The font they install on the SSRS cloud server is called Grandzebu Version 2.00 (2008) found here:  http://grandzebu.net/informatique/codbar/code128.htm

Download: [Download Code 128 Font](Reference/code128.ttf)

## How to use
SSRS reports need to be XML encoded. The text from `Code128.vb` can be dropped in to the code section of a report using the [Microsoft Report Builder](https://www.microsoft.com/en-US/download/details.aspx?id=53613). (Remove the `Module` begin and end lines)

Alternatively, you can use the reference code found in [SSRS_Report_Code.vb](Reference/SSRS_Report_Code.vb) which contains the `<code>` block of an SSRS report. 

Then textboxes where barcodes are desired can be ran through `GetCode128EncodedString` to get the text you would then display with the `Code 128` font.

Example: `=Code.GetCode128EncodedString(Fields!EmpID.Value)`
