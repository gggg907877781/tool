package tool;

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;

public class ExcelRead {
	
	
	@Test
	public void main() {
		String filePath = "g:\\Desktop\\南网企业运管指标清单V1.5-20190417.xlsx";
		ExcelRead er = new ExcelRead();
		er.getValues(filePath);
	}

	public String getValues(String filePath) {
		int a = 0;
		String values = null;
		try {
			// 创建对Excel工作簿文件的引用
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream("filePath"));
			// 创建对工作表的引用。
			// 本例是按名引用（让我们假定那张表有着缺省名"Sheet1"）
			HSSFSheet sheet = workbook.getSheet("Sheet1");
			// 也可用getSheetAt(	int index)按索引引用，
			// 在Excel文档中，第一张工作表的缺省索引是0，
			// 其语句为：HSSFSheet sheet = workbook.getSheetAt(0);
			// 读取左上端单元
			a = sheet.getLastRowNum();
			System.out.println(a);
			for (int j = 1; j <= a; j++) {
				HSSFRow row = sheet.getRow(j);
				System.out.println("-----------------------第" + j + "行数据----------------");
				for (int i = 0; i < row.getLastCellNum(); i++) {
					HSSFCell cell = row.getCell(i);
					// 输出单元内容，cell.getStringCellValue()就是取所在单元的值
					values = cell.getStringCellValue();
					System.out.println("单元格内容是： " + values);
				}
			}
		} catch (Exception e) {
			System.out.println("已运行xlRead() : " + e);
		} finally {
		}
		return values;
	}
	
	

}
