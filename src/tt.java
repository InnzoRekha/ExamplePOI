import java.io.*;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.util.DateFormatConverter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;

public class tt {

 private static void drawImageOnExcelSheet(XSSFSheet sheet, int row, int col, 
  int height, int width, int pictureIdx) throws Exception {

  CreationHelper helper = sheet.getWorkbook().getCreationHelper();

  Drawing drawing = sheet.createDrawingPatriarch();

  ClientAnchor anchor = helper.createClientAnchor();
  anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);

  anchor.setCol1(col); //first anchor determines upper left position
  anchor.setRow1(row);

  anchor.setRow2(row); //second anchor determines bottom right position
  anchor.setCol2(col);
  anchor.setDx2(Units.toEMU(width)); //dx = left + wanted width
  anchor.setDy2(Units.toEMU(height)); //dy= top + wanted height

  drawing.createPicture(anchor, pictureIdx);

 }

 public static void main(String[] args) throws Exception {
  Workbook wb = new XSSFWorkbook();
  Sheet sheet = wb.createSheet();

  InputStream is = new FileInputStream("C:\\Temp\\car1.png");
  byte[] bytes = IOUtils.toByteArray(is);
  int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
  is.close();

  String gap = "      ";
  
  Row row = null;
  for (int r = 0; r < 1000; r++ ) {
	  row=sheet.createRow(r);
	  row.createCell(1).setCellValue(gap + "Picture " + (r+1));
   //drwImageOnExcelSheet((XSSFSheet)sheet, r, 1, 12, 12, pictureIdx);
	Cell cell= row.createCell(2);
	CellStyle cellStyle = wb.createCellStyle();
	
	String excelFormatPattern = DateFormatConverter.convert(java.util.Locale.ENGLISH, "dd/mm/yyyy");
	DataFormat poiFormat = wb.createDataFormat();
	cellStyle.setDataFormat(poiFormat.getFormat(excelFormatPattern));
	cell.setCellStyle(cellStyle);
	cell.setCellValue(new Date());
	int widthUnits = 3000;
	sheet.setColumnWidth(2, widthUnits);

  }

  wb.write(new FileOutputStream("C:\\Temp\\ExcelDrawImagesOnCellLeft.xls"));
  wb.close();
 }
}