package tool;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import jxl.Sheet;
import jxl.Workbook;
import org.junit.Test;

public class ExcelRead {
	
	
	@Test
	 public void main() {
		List s = parseExcel(new File("G:\\Desktop\\南网企业运管指标清单V1.5-20190417.xlsx"));
		System.out.println(s);
    }

	
	@Test
	public void testFile () {
		
		File file = new File ("G:\\Desktop\\南网企业运管指标清单V1.5-20190417.xlsx");
		
		System.out.println(file.isFile());
	}
	public List<List<String>> parseXls(File file) {
        try {
            Workbook workbook = Workbook.getWorkbook(file);
            Sheet sheet = workbook.getSheet(1);
            List<List<String>> list = new ArrayList<List<String>>();
            for (int i = 0; i < sheet.getRows(); i++) {
                List<String> rowList = new ArrayList<String>();
                for (int j = 31; j < sheet.getColumns(); j++) {
                    rowList.add(sheet.getCell(j, i).getContents());
                }
                list.add(rowList);
            }
            // test
            for (List<String> rowList : list) {
                for (String s : rowList)
                    System.out.print(s + ",");
                System.out.println();
            }

            return list;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    public List<List<String>> parseExcel(File file) {
        try {
        	
            InputStream fis = new FileInputStream(file);
            
            String fileName = file.getName();
            
            org.apache.poi.ss.usermodel.Workbook workbook = null;
            
            if (fileName.toLowerCase().endsWith("xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else if (fileName.toLowerCase().endsWith("xls")) {
                workbook = new HSSFWorkbook(fis);
            }
            
            
            org.apache.poi.ss.usermodel.Sheet sheet = workbook.getSheetAt(1);
            
            List<List<String>> list = new ArrayList<List<String>>();
            
           // Iterator<Row> rowIterator = sheet.iterator();
            
            for (int i = 0; i < sheet.getRow; i++) {
                List<String> rowList = new ArrayList<String>();
                for (int j = 31; j < sheet.getColumns(); j++) {
                    rowList.add(sheet.getCell(j, i).getContents());
                }
                list.add(rowList);
            }
            
          /*  
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                List<String> rowList = new ArrayList<String>();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch(cell.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        rowList.add("" + cell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING:
                    default: 
                        rowList.add(cell.getStringCellValue());
                        break;
                    }
                }
                list.add(rowList);
            }
            // test
            for (List<String> rowList : list) {
                for (String s : rowList)
                    System.out.print(s + ",");
                System.out.println();
            }*/
            return list;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }

    }

   

}
