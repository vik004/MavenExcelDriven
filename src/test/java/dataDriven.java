import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class dataDriven {
	
	//Identify Testcase coloumn by scanning entire 1st row
	//Once coloumn is identified than scan enire testcase coloum to identify purchase testcase
	//after you grab purchase testcase row = pull all the data of that row and feed into test
	
	public ArrayList<String> getData(String testcasename) throws IOException {
	

	//fileInputSteam argument
	
	ArrayList<String> a= new ArrayList<String>();
	
	FileInputStream fis=new FileInputStream("D:\\ResourceStuff\\datatest.xlsx");
	XSSFWorkbook workbook=new XSSFWorkbook(fis);
	
	int sheets=workbook.getNumberOfSheets();
	for(int i=0; i<sheets;i++) 
	{
		if(workbook.getSheetName(i).equalsIgnoreCase("testdata"))
		{
		XSSFSheet sheet=workbook.getSheetAt(i);
		//Identify Testcase coloumn by scanning entire 1st row
		
		Iterator<Row> rows= sheet.rowIterator();// sheet is collection of rows
		Row firstrow=rows.next();
		Iterator<Cell> ce= firstrow.cellIterator();//row is collection of cells
		int k= 0;
		int coloumn = 0;
		while(ce.hasNext())
		{
			Cell value=ce.next();
			if(value.getStringCellValue().equalsIgnoreCase("TestCases"))
			{
				coloumn=k;
				
			}
			k++;
		}
		System.out.println(coloumn);
		
		//Once coloumn is identified than scan enire testcase coloum to identify purchase testcase
		while(rows.hasNext())
		{
			Row r= rows.next();
			if(r.getCell(coloumn).getStringCellValue().equalsIgnoreCase(testcasename))
				{
					//after you grab purchase testcase row = pull all the data of that row and feed into test
					Iterator<Cell> cv= r.cellIterator();
					while(cv.hasNext())
					{
						Cell c=cv.next();
						if(c.getCellType()==CellType.STRING)
						{
							a.add(c.getStringCellValue());
						}
						else {
							
							a.add(NumberToTextConverter.toText(c.getNumericCellValue()));
						}
					}
				}
			
		}
		
		}
	 }
		return a;
	}

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		

	}

}
