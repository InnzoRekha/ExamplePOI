import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.*;

import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;

import java.io.InputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


class CenterImageTest {

 public static void main(String[] args) {
  try {

   Workbook wb = new XSSFWorkbook();
   Sheet sheet = wb.createSheet("Sheet1");

   //create the cells A1:F1 with different widths
   Cell cell = null;
   Row row = sheet.createRow(0);
   for (int col = 0; col < 6; col++) {
    cell = row.createCell(col);
    sheet.setColumnWidth(col, (col+1)*5*256);
   }

   //merge A1:F1
   sheet.addMergedRegion(new CellRangeAddress(0,0,0,5));

   //load the picture
   InputStream inputStream = new FileInputStream("C:\\Temp\\car1.png");
   byte[] bytes = IOUtils.toByteArray(inputStream);
   int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
   inputStream.close();

   //create an anchor with upper left cell A1
   CreationHelper helper = wb.getCreationHelper();
   ClientAnchor anchor = helper.createClientAnchor();
   anchor.setCol1(0); //Column A
   anchor.setRow1(0); //Row 1

   //create a picture anchored to A1
   Drawing drawing = sheet.createDrawingPatriarch();
   Picture pict = drawing.createPicture(anchor, pictureIdx);

   //resize the pictutre to original size
   pict.resize();

   //get the picture width
   int pictWidthPx = pict.getImageDimension().width;
System.out.println(pictWidthPx);

   //get the cell width A1:F1
   float cellWidthPx = 0f;
   for (int col = 0; col < 6; col++) {
    cellWidthPx += sheet.getColumnWidthInPixels(col);
   }
System.out.println(cellWidthPx);

   //calculate the center position
   int centerPosPx = Math.round(cellWidthPx/2f - (float)pictWidthPx/2f);
System.out.println(centerPosPx);

   //determine the new first anchor column dependent of the center position 
   //and the remaining pixels as Dx
   int  anchorCol1 = 0;
   for (int col = 0; col < 6; col++) {
    if (Math.round(sheet.getColumnWidthInPixels(col)) < centerPosPx) {
     centerPosPx -= Math.round(sheet.getColumnWidthInPixels(col));
     anchorCol1 = col + 1;
    } else {
     break;
    }
   }
System.out.println(anchorCol1);
System.out.println(centerPosPx);

   //set the new upper left anchor position
   anchor.setCol1(anchorCol1);
   //set the remaining pixels up to the center position as Dx in unit EMU
   anchor.setDx1( centerPosPx * Units.EMU_PER_PIXEL);

   //resize the pictutre to original size again
   //this will determine the new bottom rigth anchor position
   pict.resize();

   FileOutputStream fileOut = new FileOutputStream("C:\\Temp\\test123.xlsx");
   wb.write(fileOut);
   fileOut.close();

  } catch (IOException ioex) {
	  System.out.println(ioex.getMessage());
	  ioex.printStackTrace();
  }
 }
}
