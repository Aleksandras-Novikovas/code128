# code128
Font and VB script for Code128 barcode generation

## Installation and usage
Install provided font.

Add functions from file Code128.vb to your Excel.
For LibreOffice you will have to add line `Option VBASupport 1` to the top of your module.

In your sheet make cell value `=Code128("Text for barcode")` and set it's font to Code128.
Make sure font size is big enough for scanner to read.

## Generated code length
Function Code128 automatically selects best subset (A, B, C) and switches between subsets
to make final result as short as possible.
Function returns empty string if provided text can not be represented with Code128 barcode.
