package FileConvertGroup.FileConvertGroupID;



import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
 
import java.io.FileReader;
import java.util.Iterator;

public class Write_JSON_Excel {

private static String[] headers = {"UserID", 	"Location",
								   "Question1", "UserAnser1", "Check1", 
								   "Question2", "UserAnser2", "Check2",
								   "Question3", "UserAnser3", "Check3", 
								   "Question4", "UserAnser4", "Check4",
								   "Question5", "UserAnser5", "Check5", 
								   "Question6", "UserAnser6", "Check6",
								  };

	public static void main(String[] args) 
	{
	
		//Blank workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		
		//Create a blank sheet
		XSSFSheet quiz1 = workbook.createSheet("Quiz 1");
		XSSFSheet quiz2 = workbook.createSheet("Quiz 2");
		XSSFSheet quiz3 = workbook.createSheet("Quiz 3");
		
		generateSheet(workbook,populate("pre_quiz_0", "player.json"),quiz1,headers);
		generateSheet(workbook,populate("pre_quiz_1", "player.json"),quiz2,headers);
		generateSheet(workbook,populate("pre_quiz_2", "player.json"),quiz3,headers);
	
	}
	
	public static void generateSheet(XSSFWorkbook workbook, Map<Integer, Object[]> columns, XSSFSheet sheet, String[] headers) {
		Map<Integer, Object[]> data = columns;
		
		
		//Iterate over data and write to sheet
		Set<Integer> keyset = data.keySet();
		int rownum = 1;
		
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontHeightInPoints((short) 14);
		headerFont.setColor(IndexedColors.AUTOMATIC.getIndex());
	
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setFont(headerFont);
		
		// Create a Header Row
		Row headerRow = sheet.createRow(0);
		
		// Create cells
		for(int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(headerCellStyle);
		}
		
		for (Integer key : keyset)
		{
			Row row = sheet.createRow(rownum++);
			Object [] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr)
			{
				Cell cell = row.createCell(cellnum++);
				if(obj instanceof String)
				cell.setCellValue((String)obj);
				else if(obj instanceof Integer)
				cell.setCellValue((Integer)obj);
			}
		}
		try
		{
			FileOutputStream out = new FileOutputStream(new File("AC-GS-Reports.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("**Tab titled "+sheet.getSheetName() +" has been Modified*!**");
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}
	
	@SuppressWarnings("unchecked")
	public static Map<Integer, Object[]> populate(String pattern, String filename) {
		Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();	
		JSONParser parser = new JSONParser();
		try {
			Object obj = parser.parse(new FileReader(filename));
			JSONArray playerdata = (JSONArray) obj;
			Iterator<JSONObject> listOfPlayerData=playerdata.iterator();
			int record=1;
			while(listOfPlayerData.hasNext()) {
				JSONObject currentUser=listOfPlayerData.next();
				Object[] column=new Object[100];
				int columnPosition=0;		
				JSONObject question = ((JSONObject)currentUser.get("privateData"));
				boolean found=false;
				if(!question.isEmpty()) {
					Map<Object, Object> arr=question;
					for (Map.Entry<Object,Object> entry : arr.entrySet())   
					{  
						if(entry.getKey().toString().startsWith(pattern)) {
							
							Map<Object, Object> currentQuestion=(JSONObject)entry.getValue();
							for(Map.Entry<Object,Object> e : currentQuestion.entrySet()){
								found=true;
								if(columnPosition==0) {
								column[columnPosition++]=currentUser.get("userName").toString();
								column[columnPosition++]=((JSONObject)currentUser.get("location")).get("city").toString();
								}
								column[columnPosition++]=((JSONObject)e.getValue()).get("question").toString();
								column[columnPosition++]=((JSONObject)e.getValue()).get("playerAnswer").toString();
								column[columnPosition++]=((JSONObject)e.getValue()).get("correct").toString();
							}			
						}			  
					} 
					if(found)
						data.put(record++, column);
				}		
			}	
		} catch (Exception e) {
			e.printStackTrace();
		}
		return data;
	}

}

