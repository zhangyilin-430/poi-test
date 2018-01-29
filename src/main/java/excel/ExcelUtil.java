package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtil {

	private ExcelUtil() {
	};

	private static  List<String>  columns = Arrays.asList("父组件名称", "父组件代号 ");// 要解析excel中的列名
	private static int sheetNum = 0;// 要解析的sheet下标
	private static StringBuilder retJson = new StringBuilder();// 拼接语句

	/**
	 * poi读取excle
	 * 
	 * @return
	 */
	public static String readExcel(File file) {

		InputStream inStream = null;
		try {
			// 读取文件
			inStream = new FileInputStream(file);
			// 文件转化为HSSFWorkbook=Excel2003 扩展名是.xls
			HSSFWorkbook workbook = new HSSFWorkbook(inStream);
			HSSFSheet sheet = null;
			// 遍历sheet
			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
				sheet = workbook.getSheetAt(i);// 获取下标为i的sheet页
				retJson.append("[");
				readExcelSheet(sheet); // 读取sheet页
				if (i < workbook.getNumberOfSheets())// 使用，隔离每行
					retJson.append(",");
				retJson.append("]");
				// System.out.println("---Sheet表"+i+"处理完毕---");
			}
		} catch (Exception e) {
			try {
				// 读取文件
				inStream = new FileInputStream(file);
				// 文件转化为XSSFWorkbook=Excel2007 扩展名是.xlsx
				XSSFWorkbook workbook = new XSSFWorkbook(inStream);
				workbook.getNumberOfSheets();
				XSSFSheet sheet = workbook.getSheetAt(sheetNum);
				int lastRowNum = sheet.getLastRowNum();// 最后一行
				retJson.append("[");
				for (int i = 0; i < lastRowNum; i++) {
					XSSFRow row = sheet.getRow(i);// 获得行
					String rowJson = readExcelRow(row);
					retJson.append(rowJson);
					if (i < lastRowNum - 1)
						retJson.append(",");
				}
				retJson.append("]");
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		} finally {
			close(null, inStream);
		}
		return retJson.toString();
	}

	/**
	 * 读取sheet页
	 * 
	 * @param sheet
	 * @return
	 */
	private static void readExcelSheet(HSSFSheet sheet) {
		/***************************/
		// getPhysicalNumberOfRows() 获取的是物理行数，也就是不包括那些空行（隔行）的情况。
		// getLastRowNum() 获取的是最后一行的编号（编号从0开始）。
		// 不过POI里似乎没有没有这样的方法，getNextPhysicalRow()。
		// 所以只好从getFirstRowNum()到getLastRowNum()遍历，如果null==currentRow，验证下一行。
		/**************************/
		int lastRowNum = sheet.getPhysicalNumberOfRows();// 最后一行
		retJson.append("[");
		for (int i = 0; i < lastRowNum; i++) {// 获取每行
			HSSFRow row = sheet.getRow(i);// 获得行
			String rowJson = readExcelRow(row);// 获取单元格
			retJson.append(rowJson);
			if (i < lastRowNum - 1)// 使用，隔离每行
				retJson.append(",");
		}
		retJson.append("]");
	}

	/**
	 * 读取sheet页
	 * 
	 * @param sheet
	 * @return
	 */
	private static void readExcelSheet(XSSFSheet sheet) {

		/***************************/
		// getPhysicalNumberOfRows() 获取的是物理行数，也就是不包括那些空行（隔行）的情况。
		// getLastRowNum() 获取的是最后一行的编号（编号从0开始）。
		// 不过POI里似乎没有没有这样的方法，getNextPhysicalRow()。
		// 所以只好从getFirstRowNum()到getLastRowNum()遍历，如果null==currentRow，验证下一行。
		/**************************/
		int lastRowNum = sheet.getPhysicalNumberOfRows();// 最后一行
		retJson.append("[");
		for (int i = 0; i < lastRowNum; i++) {// 获取每行
			XSSFRow row = sheet.getRow(i);// 获得行
			String rowJson = readExcelRow(row);
			retJson.append(rowJson);
			if (i < lastRowNum - 1)
				retJson.append(",");
		}
		retJson.append("]");
	}

	/**
	 * 读取行值
	 * 
	 * @return
	 */
	private static String readExcelRow(HSSFRow row) {
		StringBuilder rowJson = new StringBuilder();
		int lastCellNum = ExcelUtil.columns.size();// 最后一个单元格
		rowJson.append("{");
		for (int i = 0; i < lastCellNum; i++) {
			HSSFCell cell = row.getCell(i);
			String cellVal = readCellValue(cell);
			rowJson.append(toJsonItem(columns.get(i), cellVal));
			if (i < lastCellNum - 1)
				rowJson.append(",");
		}
		rowJson.append("}");
		return rowJson.toString();
	}

	/**
	 * 读取行值
	 * 
	 * @return
	 */
	private static String readExcelRow(XSSFRow row) {
		StringBuilder rowJson = new StringBuilder();
		int lastCellNum = ExcelUtil.columns.size();// 最后一个单元格
		rowJson.append("{");
		for (int i = 0; i < lastCellNum; i++) {
			XSSFCell cell = row.getCell(i);
			String cellVal = readCellValue(cell);
			rowJson.append(toJsonItem(columns.get(i), cellVal));
			if (i < lastCellNum - 1)
				rowJson.append(",");
		}
		rowJson.append("}");
		return rowJson.toString();
	}

	/**
	 * 读取单元格的值
	 * 
	 * @param hssfCell
	 * @return
	 */
	@SuppressWarnings("static-access")
	private static String readCellValue(HSSFCell hssfCell) {
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(hssfCell.getNumericCellValue());
		} else {
			return String.valueOf(hssfCell.getRichStringCellValue());
		}
	}

	/**
	 * 读取单元格的值
	 * 
	 * @param hssfCell
	 * @return
	 */
	@SuppressWarnings("static-access")
	private static String readCellValue(XSSFCell hssfCell) {
		if (hssfCell.getCellType() == hssfCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == hssfCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(hssfCell.getNumericCellValue());
		} else {
			return String.valueOf(hssfCell.getRichStringCellValue());
		}
	}

	/**
	 * 转换为json对
	 * 
	 * @return
	 */
	private static String toJsonItem(String name, String val) {
		return "\"" + name + "\":\"" + val + "\"";
	}

	/**
	 * 关闭io流
	 * 
	 * @param fos
	 * @param fis
	 */
	private static void close(OutputStream out, InputStream in) {
		if (in != null) {
			try {
				in.close();
			} catch (IOException e) {
				System.out.println("InputStream关闭失败");
				e.printStackTrace();
			}
		}
		if (out != null) {
			try {
				out.close();
			} catch (IOException e) {
				System.out.println("OutputStream关闭失败");
				e.printStackTrace();
			}
		}
	}

	public static List<String> getColumns() {
		return ExcelUtil.columns;
	}

	public static void setColumns(List<String> columns) {
		ExcelUtil.columns = columns;
	}

	public static int getSheetNum() {
		return sheetNum;
	}

	public static void setSheetNum(int sheetNum) {
		ExcelUtil.sheetNum = sheetNum;
	}

	public static void main(String[] args) {
		readExcel(new File("C:\\Users\\Administrator\\Desktop\\BOM-WXB903 - 副本.xlsx"));
	}

}