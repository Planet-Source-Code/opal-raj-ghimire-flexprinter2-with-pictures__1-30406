Read this B4 you use.

While using MSHFlex Grid you must set colwidth to each column. 
See load event of the form, colsetupcode method also generates the 
code needed but you must slide every column.

If you issued "Flex.ColWidth(3) = 0 " column#3 does not print.
If you issued "Flex.RowHeight(3) = 0 " row#3 does not print.

If you try to make rowheight smaller than text, it looks ok in 
picture box but while printing it prints over the cell one below that.
If there is three dots at the end of text, it prints ok(turnicated) 
Otherwise it prints over the below cell.

Value of RowFrom should always be smaller than RowTo.
CurX and CurY are like CurrentX ans CurrentY of Vb.

Printing is quite nice in my laserjet 6L. 

Pictures: Only BMP picture allowed. If you use other types like icon,
an error occures.
PixHeight and PixWidth must be set for pictures that size other than
32X32 pixels.

If you encounter total blackness in output, this happens because the color
value returend from cellbackcolor or backcolorfixed of MS(H) flex
grid control are invalid. First set them and try again. It happens when 
one uses default values.

I am just an ordinary one, these codes may still
contain error. If you find one inform me I will try to fix it.

Do send your comments, suggestions and improvement.

Opal Raj Ghimire
http://geocities.com/opalraj/vb 
Kathmandu, Nepal
2002, Jan