package solution;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
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
		XSSFSheet fromSh =  wb.getSheet("list1");
		XSSFRow firstRow = fromSh.getRow(0);//строка заголовка с инструкциями
		XSSFWorkbook resultWb = new XSSFWorkbook();
		XSSFSheet resultSh = resultWb.createSheet();
		XSSFRow toRow = resultSh.createRow(0);//строка на которую переносятся данные	
		int j = 0;	
		copyHeader(firstRow, toRow); 
		XSSFRow fromRow; //активная строки (будет помещена в результат, а дубли будут уничтожены)
		int toRowInd = 1; //индекс строки в результирующей таблице
		int curRow = 1;//индекс строки в оригинальной таблице
		while(fromSh.getRow(curRow) != null) {//перебор строк
			fromRow = fromSh.getRow(curRow);//строка с которой переносятся данные
			int i = curRow;			
			while(fromSh.getRow(i) != null) {//сравнение строк с текущей
				j = 0;	
				while(fromSh.getRow(0).getCell(j) != null) 	//поиск всех столбцов с критерием
				{		
					if(fromSh.getRow(0).getCell(j).toString().compareTo("Crit") == 0) {
						if(fromRow.getCell(j).toString().compareTo(fromSh.getRow(i).getCell(j).toString()) != 0 )
						{	
							break; 
						}
					}	
					j++;
				}				
				if(j == getNumberOfCells(fromSh.getRow(i)) && fromSh.getRow(i).getCell(getNumberOfCells(fromSh.getRow(i))+2) == null) { //проверка что критерии совпадают
					//и что строка ещё не брабатывалась
					// орбработанная строка помечается в колонке getNumberOfCells(fromSh.getRow(i))+2
					if(i == curRow) {
						toRow = resultSh.createRow(toRowInd);
						toRowInd++;
					}
					int k = 0;//индекс текущего столбца в оригинальной таблице
					int toK = 0;//индекс текущего столбца в результирующей таблице
					 while(fromRow.getCell(k) != null && destroySpaces(fromRow.getCell(k).toString()).length() > 0) //перебор  столбцов 
					 { 						 
						 if(firstRow.getCell(k).toString().compareTo("Con") != 0 && //если ячейка подлежит конкатенации или является критерием проверять её не нужно
						    firstRow.getCell(k).toString().compareTo("-") != 0 &&
						    firstRow.getCell(k).toString().compareTo("Crit") != 0) {
							 try {
								 getDouble(fromSh.getRow(i).getCell(k)); //попытка получить значение в виде числа
							 }
							 catch(Exception  e) {
								 System.out.print("Ошибка по адресу: " + i + " " + k);//вывод ошибки ввода
								 System.out.println();
								 throw new Exception("Неправильные данные");
							 }
						 } 
						 if(toRow.getCell(toK) == null && firstRow.getCell(k).toString().compareTo("-") != 0) { //заполнение ячейки начальными данными
							 String value = String.valueOf(destroySpaces(fromSh.getRow(i).getCell(k).toString())).replaceFirst("\\.0+$", "");//удаление .0 у чиловых значений
							 toRow.createCell(toK).setCellValue(value);							 
						 }else	{
							 if(firstRow.getCell(k).toString().compareTo("Con") == 0)	 {//конкатинация строк 
								 String oldValue =  destroySpaces(toRow.getCell(toK).toString());
								 oldValue = String.valueOf(oldValue).replaceFirst("\\.0+$", "");
								 String newValue =  destroySpaces(fromSh.getRow(i).getCell(k).toString());
								 newValue = String.valueOf(newValue).replaceFirst("\\.0+$", "");
								 toRow.getCell(toK).setCellValue(oldValue + newValue);
							 }
						  if(firstRow.getCell(k).toString().compareTo("Max") == 0) {//поиск максимума 
							 double num1 = getDouble(toRow.getCell(toK)); 
							  double num2 =  getDouble(fromSh.getRow(i).getCell(k));
							  if(num2 > num1) {
								  toRow.getCell(toK).setCellValue(num2); 
							  } 							  
						  }
						  if(firstRow.getCell(k).toString().compareTo("Min") == 0) {//поиск минимума
							  double  num1 = getDouble(toRow.getCell(toK));
							  double num2 =  getDouble(fromSh.getRow(i).getCell(k));
							  if(num2 < num1) {
								  toRow.getCell(toK).setCellValue(num2);
							  } 
						  }
						  if(firstRow.getCell(k).toString().compareTo("Sum") == 0) {//поиск суммы							  
							  double num1 = getDouble(toRow.getCell(toK));
							  double num2 =  getDouble(fromSh.getRow(i).getCell(k));
							  toRow.getCell(toK).setCellValue(num2+num1);
						  }
						 }
						 
						 if(firstRow.getCell(k).toString().compareTo("-") != 0) {
							 toK++;
						 }						 
					  k++; 					  
					  }
					 fromSh.getRow(i).createCell(getNumberOfCells(fromSh.getRow(i))+2).setCellValue("+");//пометка строки как обработанной
				}
				
				i++;						
			}
			curRow++;
		}		
		FileOutputStream fileOut  = new FileOutputStream("./result.xlsx");
		resultWb.write(fileOut);
		 System.out.println("Программа выполнена");
	}
	private static String destroySpaces(String str) {//удаление лишних прпобелов в строке
		StringBuilder newStr = new StringBuilder(str);
		
		while(newStr.charAt(0) == ' ') {
			newStr.deleteCharAt(0);
		}
		while(newStr.charAt(newStr.length()-1)== ' ') {
			
			newStr.deleteCharAt(newStr.length()-1);
			
		}
		return newStr.toString();
	}
	
	private static int getNumberOfCells(XSSFRow row) {// пересцёт кол-ва ячеек в строке
		int result = 0;
		
		while (row.getCell(result) != null && destroySpaces(row.getCell(result).toString()).length() > 0) {
			result++;			
		}
		return result;
	}
	private static double getDouble(XSSFCell cell) throws Exception{ //получене числового значения из ячейки
		
		double result = 0;
		DataFormatter formatter = new DataFormatter();
		String str = formatter.formatCellValue(cell);
		result = Double.parseDouble(destroySpaces(str) );
		
		return result;
	}
	private static void  copyHeader(XSSFRow rowFrom, XSSFRow rowTo) throws Exception { //проверка заголовка на корректность и удаление сьолюцов помеченных "-"
		int j = 0;
		int toJ = 0;
		while(rowFrom.getCell(j) != null) {
			if(!((rowFrom.getCell(j).toString().compareTo("-") == 0) ||
			   (rowFrom.getCell(j).toString().compareTo("Min") == 0) ||
			   (rowFrom.getCell(j).toString().compareTo("Max") == 0) ||
			   (rowFrom.getCell(j).toString().compareTo("Con") == 0) ||
			   (rowFrom.getCell(j).toString().compareTo("Sum") == 0) ||
			   (rowFrom.getCell(j).toString().compareTo("Crit") == 0))){
				System.out.println("Неправильно задан заголовок. В заголовке допустимы только -, Max, Min, Con, Crit.");
				throw new Exception("Ошибка в заголовке");
			}
			if(rowFrom.getCell(j).toString().compareTo("-") != 0) {
				rowTo.createCell(toJ).setCellValue(rowFrom.getCell(j).toString());
				toJ++;
			}
			j++;
		}
	}
}




