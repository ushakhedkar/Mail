import java.io.IOException;

public class TestConfig
{	
	ReadExcel excel;

public void Excel() throws IOException
{
	excel=new ReadExcel("C:\\Users\\admin\\Documents\\mail.xlsx","Mail");
	int row=excel.getRowCount();
	int cols=excel.getCellCount();
	System.out.println(excel.getData(row, cols));	
	String str=excel.getData(row, cols);
	System.out.println(str);
}
	
	//https://myaccount.google.com/lesssecureapps
		public static String server="smtp.gmail.com";
		public static String from = "ushakhedkaruck@gmail.com"; //enter your username
		public static String password = "khedkar@123"; //enter your password
		public static String[] to ={"ushakhedkaruck@gmail.com"} ;
		public static String subject = "Test Report";
		
		public static String messageBody ="TestMessage";
	/*	public static String attachmentPath="G:\\Google Language Screenshots\\Google.jpeg";
		public static String attachmentName="G:\\error.jpg";
		{"ushakhedkaruck@gmail.com"} ;	*/			
}
