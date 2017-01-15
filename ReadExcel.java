import java.io.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadExcel 
{
	static int numSheet = 1;
	
	public static void main(String[] args) throws IOException, FileNotFoundException
	{
		FileInputStream excelFile = new FileInputStream("Book1.xls");
		POIFSFileSystem eFile = new POIFSFileSystem(excelFile); 
		HSSFWorkbook eWorkbook = new HSSFWorkbook(eFile);
		int numSheets = eWorkbook.getNumberOfSheets();
		
		for(int i = 0; i < numSheets; i++)
		{
			HSSFSheet eSheet = eWorkbook.getSheetAt(i);
			readSheet(eSheet);
		}
		
		eWorkbook.close();
		excelFile.close();
	}
	
	public static void readSheet(HSSFSheet sheet) 
	{
		int numRows = sheet.getPhysicalNumberOfRows();
		int numCol;
		HSSFRow rows;
		HSSFCell cells;
		if(numRows != 0)
		{
			System.out.println("SHEET " + numSheet + ":");
			for(int i = 0; i < numRows; i++)
			{
				rows = sheet.getRow(i);
				numCol = rows.getPhysicalNumberOfCells();
				for(int j = 0; j < numCol; j++)
				{
					cells = rows.getCell(j);
					System.out.format("%30s", cells);
				}
				System.out.println();
			}
		numSheet++;
		}
	}

}
