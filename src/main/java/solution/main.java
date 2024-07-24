package solution;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class main {
	

	public static void main(String[] args) throws Exception{
		XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("./data1.xlsx"));
		XSSFSheet sh =  wb.getSheet("list1");
		XSSFRow firstRow = sh.getRow(0);//строка заголовка с инструкциями
		int secondRowInd = 1; //индекс текущей активной строки
		XSSFRow secondRow; //активная строки (будет помещена в результат, а дубли будут уничтожены)
		int j = 0;	
		
		while(firstRow.getCell(j) != null) 	//цикл удаляет колонки не помеченные к отображению	
		{
			if(firstRow.getCell(j).toString().charAt(0) == '-' ) {
				deleteColumn(sh, j);
				j--;
			}
			j++;
		}
		
		while(sh.getRow(secondRowInd) != null) {//перебор всех активных строк
			 secondRow = sh.getRow(secondRowInd);//установка активной обрабатываемой строки
			int i = secondRowInd + 1;
			while(sh.getRow(i) != null) {// перебор неактивных строк				
				j = 0;	
				while(firstRow.getCell(j) != null && destroySpaces(firstRow.getCell(j).toString()).length() > 0) 	//перебор столбцов	
				{	
					if(firstRow.getCell(j).toString().charAt(0) == 'E') {//конкатинация строк
						String oldValue = sh.getRow(secondRowInd).getCell(j).toString();
						String newValue = sh.getRow(i).getCell(j).toString();
						sh.getRow(secondRowInd).getCell(j).setCellValue(oldValue + newValue);
					}
					if(firstRow.getCell(j).toString().charAt(0) == 'C') {//поиск максимума
						double num1 = getDouble(sh.getRow(secondRowInd).getCell(j));
						double num2 = getDouble(sh.getRow(i).getCell(j));
						if(num2 > num1) {
							sh.getRow(secondRowInd).getCell(j).setCellValue(num2);
						}
					}
					if(firstRow.getCell(j).toString().charAt(0) == 'D') {//поиск минимума
						double num1 = getDouble(sh.getRow(secondRowInd).getCell(j));
						double num2 = getDouble(sh.getRow(i).getCell(j));
						if(num2 < num1) {
							sh.getRow(secondRowInd).getCell(j).setCellValue(num2);
						}
					}
					if(firstRow.getCell(j).toString().charAt(0) == 'B') {//поиск суммы
						double num1 = getDouble(sh.getRow(secondRowInd).getCell(j));
						double num2 = getDouble(sh.getRow(i).getCell(j));
						sh.getRow(secondRowInd).getCell(j).setCellValue(num1+ num2);
					}
					if(firstRow.getCell(j).toString().charAt(0) == 'A') {
						if(!(secondRow.getCell(j).toString() == sh.getRow(i).getCell(j).toString()))
						{	
							break; //выход из цикла если строка должна исчезнуть в ходе слияния 
						}
					}		
					
					j++;
				}
				if(j == getNumberOfCells(sh.getRow(i))) { 
					
					removeRow(sh, i);//удаление лишей строки
				}else {
					i++;	
				}				
			}
			secondRowInd++;
		}		
		FileOutputStream fileOut  = new FileOutputStream("./result.xlsx");
		wb.write(fileOut);
		 System.out.println("Программа выполнена");
	}
	
	public static void removeRow(XSSFSheet sheet, int rowIndex) {//функция удаления строки
	    int lastRowNum=sheet.getLastRowNum();
	    if(rowIndex>=0&&rowIndex<lastRowNum){
	        sheet.shiftRows(rowIndex+1,lastRowNum, -1);
	    }
	    if(rowIndex==lastRowNum){
	        XSSFRow removingRow=sheet.getRow(rowIndex);
	        if(removingRow!=null){
	            sheet.removeRow(removingRow);
	        }
	    }
	}
	
	private static void deleteColumn(Sheet sheet, int columnToDelete) { //функция удаления столбца
		for (int rId = 0; rId <= sheet.getLastRowNum(); rId++) {
	        Row row = sheet.getRow(rId);
	        for (int cID = columnToDelete; cID < row.getLastCellNum(); cID++) {
	            Cell cOld = row.getCell(cID);
	            if (cOld != null) {
	                row.removeCell(cOld);
	            }
	            Cell cNext = row.getCell(cID + 1);
	            if (cNext != null) {
	                Cell cNew = row.createCell(cID, cNext.getCellType());
	                cloneCell(cNew, cNext);
	                if(rId == 0) {
	                    sheet.setColumnWidth(cID, sheet.getColumnWidth(cID + 1));

	                }
	            }
	        }
	    }
	}

	private static void cloneCell(Cell cNew, Cell cOld) { //клонирование ячейки
	    cNew.setCellComment(cOld.getCellComment());
	    cNew.setCellStyle(cOld.getCellStyle());

	    if (CellType.BOOLEAN == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getBooleanCellValue());
	    } else if (CellType.NUMERIC == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getNumericCellValue());
	    } else if (CellType.STRING == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getStringCellValue());
	    } else if (CellType.ERROR == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getErrorCellValue());
	    } else if (CellType.FORMULA == cNew.getCellType()) {
	        cNew.setCellValue(cOld.getCellFormula());
	    }
	}
	private static String destroySpaces(String str) {//удаление лишних прпобелов в строке
		StringBuilder newStr = new StringBuilder(str);
		
		for(int i = str.length()-1; i >= 0; i--) {
			if(!isLetterOrDigit(str.charAt(i))) {
				newStr.deleteCharAt(i);
			}
		}
		return newStr.toString();
	}
	private static boolean isLetterOrDigit(char c) {//проверка на то что символ буква или цифра
		return (c >= 'a' && c <= 'z') ||
		           (c >= 'A' && c <= 'Z') ||
		           (c >= '0' && c <= '9');
	}
	private static int getNumberOfCells(XSSFRow row) {// пересцёи кол-ва ячеек в строке
		int result = 0;
		
		while (row.getCell(result) != null && destroySpaces(row.getCell(result).toString()).length() > 0) {
			result++;			
		}
		return result;
	}
	private static double getDouble(XSSFCell cell) { //получене числового значения из ячейки
		DataFormatter formatter = new DataFormatter();
		String str = formatter.formatCellValue(cell);
		double result = Double.parseDouble(destroySpaces(str) );
		return result;
	}
}




