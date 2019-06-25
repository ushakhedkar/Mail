
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ReadExcel 
{
	Workbook book;
	Sheet sheet;
	
	public ReadExcel(String fname,String sheetname) throws  IOException    
	{
		File f=new File(fname);
		FileInputStream fis=new FileInputStream(f);
		book=WorkbookFactory.create(fis);
		sheet=book.getSheet(sheetname);
	
	}

		int rowcount=0;
		int cellcount=0;
	
	public int getRowCount()
	{		
		for(Row row:sheet)
			
		{
			rowcount++;
		}
		return rowcount;
	}
	public int getCellCount()
	{		
		for(Row row1:sheet)		
		{
			for(Cell cell:row1)
			{
				cellcount++;
			}
			break;
		}
		return cellcount;
	}
	Row row;
	Cell cell;
	String email =null;

	public String getData(int rw,int cols)
	{		
        for (int r = 0; r <=rw; r++) 
        {
            row = sheet.getRow(r); // bring row
            if (row != null) {
                for (int c = 0; c <=cols; c++) 
                {
                    cell = row.getCell(c);
                    if (cell != null) 
                    {
                        switch (cell.getCellType()) 
                        {                   

                        case FORMULA:
                        	String form=(String)cell.getCellFormula();
                            email = String.valueOf(form);
                            break;

                        case NUMERIC:
                        	double d=cell.getNumericCellValue();
                            email = String.valueOf(d);
                            break;

                        case STRING:
                            email =cell.getStringCellValue();
                            break;

                        case BLANK:
                            email = "[BLANK]";
                            break;
						case ERROR:
                            email = "" + cell.getErrorCellValue();
                            break;
                        default:    
                        	System.out.println(email);
                        }                               
                    } 
                } 
            }
        } 
		
		return email;
    }
		/*ReadExcel excel=new ReadExcel("C:\\Users\\admin\\Documents","Mail");
		int row=excel.getRowCount();
		int cols=excel.getCellCount();*/

}
