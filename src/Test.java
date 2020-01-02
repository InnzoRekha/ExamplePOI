import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.util.IOUtils;

import java.io.InputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;



class Test {

 public static void main(String[] args) {
  try {
   String names[]={"rekha","senthil","varun","kumar","chitra","bablu"};
   String filenames[]={"car1.png","car2.png","car3.png","car4.png","car5.png","car6.png"};
   Workbook wb = new XSSFWorkbook();
   Sheet sheet = wb.createSheet("My Sample Excel");
   //FileInputStream obtains input bytes from the image file
   InputStream inputStream = new FileInputStream("C:\\Temp\\"+ filenames[0]);
   //Get the contents of an InputStream as a byte[].
   byte[] bytes = IOUtils.toByteArray(inputStream);
   //Adds a picture to the workbook
   int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
   InputStream inputStream1 = new FileInputStream("C:\\Temp\\"+ filenames[1]);
   //Get the contents of an InputStream as a byte[].
   byte[] bytes1 = IOUtils.toByteArray(inputStream1);
   //Adds a picture to the workbook
   int pictureIdx1 = wb.addPicture(bytes1, Workbook.PICTURE_TYPE_PNG);

   //close the input stream
   inputStream.close();
   //Returns an object that handles instantiating concrete classes
   CreationHelper helper = wb.getCreationHelper();
   //Creates the top-level drawing patriarch.
   Drawing drawing = sheet.createDrawingPatriarch();

   //Create an anchor that is attached to the worksheet
   ClientAnchor anchor = helper.createClientAnchor();

   //create an anchor with upper left cell _and_ bottom right cell
   anchor.setCol1(1); //Column B
   anchor.setRow1(2); //Row 3
   anchor.setCol2(2); //Column C
   anchor.setRow2(3); //Row 4
   
   //Creates a picture
 //  Picture pict = drawing.createPicture(anchor, pictureIdx);

   //Reset the image to the original size
   //pict.resize(); //don't do that. Let the anchor resize the image!

   //Create the Cell B3
   for(int i=0;i<6;i++)
   {
   Row row = sheet.createRow(i);
   row.setHeight((short) 2500);
   Cell cell = row.createCell(0);
   
   sheet.setColumnWidth(0, (1+1)*5*256);
   
   cell.setCellValue(names[i]);
   Cell cell1 = row.createCell(1);
   
  // sheet.setColumnWidth(1, (1+1)*5*256);
   
   InputStream inputStream2= new FileInputStream("C:\\Temp\\"+ filenames[i]);
   System.out.println("fffff"+i);
   System.out.println(filenames[i]);
   //Get the contents of an InputStream as a byte[].
   byte[] bytes2 = IOUtils.toByteArray(inputStream2);
   //Adds a picture to the workbook
   int pictureIdx2 = wb.addPicture(bytes2, Workbook.PICTURE_TYPE_PNG);
   ClientAnchor anchor1 = helper.createClientAnchor();

   //create an anchor with upper left cell _and_ bottom right cell
   anchor1.setCol1(1); //Column B
   anchor1.setRow1(i); //Row 3
  anchor1.setCol2(2); //Column C
  anchor1.setRow2(i+1); //Row 4
 
   anchor1.setAnchorType(AnchorType.MOVE_AND_RESIZE);
   
   Picture pict1,pict2;
   
	   pict1 =  drawing.createPicture(anchor1, pictureIdx2);
   
	   sheet.setColumnWidth(1, (int)pict1.getImageDimension().getWidth() * 2 );	
	   //sheet.setRowHeight(1, (int) pict1.getImageDimension().getHeight()); //height equals to picture height
	 //Set the width of the first and second column
	   //sheet.setColumnWidth(0, pict1.getImageDimension().getWidth());
	 
   Cell cell2 = row.createCell(2);
 
   cell2.setCellValue(i);
   Cell cell3 = row.createCell(3);
 
   double x = Math.random();
   cell3.setCellValue(x);
   }
   //set width to n character widths = count characters * 256
   //int widthUnits = 20*256;
   //sheet.setColumnWidth(1, widthUnits);

   //set height to n points in twips = n * 20
   //short heightUnits = 60*20;
   //cell.getRow().setHeight(heightUnits);

   //Write the Excel file
   FileOutputStream fileOut = null;
   fileOut = new FileOutputStream("C:\\Temp\\test11.xlsx");
   wb.write(fileOut);
   fileOut.close();

  } catch (IOException ioex) {
	  System.out.println(ioex.getMessage());
	  System.out.println(ioex.getMessage());
  }
 }
}