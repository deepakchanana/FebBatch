package pkg1;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelRead
{
	
	public void excelWriteData() throws IOException, RowsExceededException, WriteException
	{
		System.out.println("changes done by deepak");
		File f=new File("../ExcelFileHandling/data2.xls");
		WritableWorkbook wk=Workbook.createWorkbook(f);
		WritableSheet ws=wk.createSheet("Deepak", 0);
		for(int i=0;i<3;i=i+1)   // loop for rows
		{
			for(int j=0;j<3;j=j+1)  // loop for columns
			{
				Label L=new Label(j, i, "DC");
				ws.addCell(L);
			}
		}
		wk.write(); // for saving the file
		wk.close(); // for closing the file
		
 	}
	
	
	public void excelReadData() throws BiffException, IOException
	{
	File f=new File("../ExcelFileHandling/data.xls");
	Workbook wk=Workbook.getWorkbook(f);
	Sheet ws=wk.getSheet(0);
	int row=ws.getRows();
	int column=ws.getColumns();
	for(int i=0;i<row;i=i+1)  // loop for rows
	{
		for(int j=0;j<column;j=j+1)  // loop for columns
		{
		Cell c1=ws.getCell(j, i);	
		System.out.println(c1.getContents());
		}
	}
	}
	public static void main(String[] args) throws BiffException, IOException, RowsExceededException, WriteException 
	{
	ExcelRead ex=new ExcelRead();
	ex.excelWriteData();
	}

}
