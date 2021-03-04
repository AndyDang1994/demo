package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;

public class ReadFromExcel {


	public String filepath = null;
	public String fileType;
	public Object[][] data = null;
	final DataFormatter df = new DataFormatter();
	public static Gson gson = new GsonBuilder().setDateFormat("dd-MM-yyyy").setPrettyPrinting().create();
	static SimpleDateFormat DtFormat = new SimpleDateFormat("dd/MM/yyyy");
	 

	/**
	 * 
	 * @param ar
	 */
	@SuppressWarnings("rawtypes")
	public static void main(String ar[]) {
		try {
			ReadFromExcel rw = new ReadFromExcel("/Users/dangkimhhoang/Documents/DCMS/AGENT_OTHER_MOVEMENT_UPLOAD_THEMPLATE.xlsx");
			WriteToExcel wr = new WriteToExcel("/Users/dangkimhhoang/Documents/DCMS/AGENT_OTHER_MOVEMENT_DOWLOAD_THEMPLATE.xlsx");
			
			model md = new model();
			md.setEmail("ahah");
			md.setName("hihi");
			md.setPassword2("heheh");
			
			List<model> lsmd = new ArrayList<model>();
			List<String> lsheader = new ArrayList<String>();
			lsheader.add("Agent Code");
			lsheader.add("Agent Movement Type");
			lsheader.add("Agent Movement Segment");
			lsheader.add("Agent Code");
			lsheader.add("Agent Code");
			lsheader.add("Agent Code");
			lsheader.add("Agent Code");
			lsmd.add(md);
			wr.writeDataToExcel().writeHeader(lsheader).writeData(lsmd).write();
			
			List<model> lsObject = new ArrayList<>();

			ArrayList<Map> lsArrMapObject = rw.readDataFromExcel();
			if (lsArrMapObject != null && lsArrMapObject.size() > 0) {
				for (Map mapData : lsArrMapObject) {
					lsObject.add(gson.fromJson(gson.toJson(mapData), model.class));
					System.out.println(gson.toJson(mapData));
				}
			}

			// Final class
			if (lsObject != null && lsObject.size() > 0) {
				for (Object obj : lsObject) {
					System.out.println(obj.toString());
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	

	/**
	 * 
	 * @param filepath
	 */
	public ReadFromExcel(String filepath) {
		if (filepath != null) {
			if (filepath.endsWith("xlsx")) {
				fileType = "xlsx";
			} else if (filepath.endsWith("xls")) {
				fileType = "xls";
			}
		}
		this.filepath = filepath;
	}

	/**
	 * Read Data From Excel
	 * 
	 * @return
	 */
	@SuppressWarnings("rawtypes")
	public ArrayList<Map> readDataFromExcel() {
		ArrayList<Map> lsData = new ArrayList<>();
		try {
			FileInputStream file = new FileInputStream(getFile());
			Workbook workbook = null;
			if (fileType.equalsIgnoreCase("xlsx")) {
				workbook = new XSSFWorkbook(file);
			} else if (fileType.equalsIgnoreCase("xls")) {
				workbook = new HSSFWorkbook(file);
			}
			int sheetIndex = 0;
			lsData = getMappedValues(workbook, sheetIndex, new model());
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return lsData;
	}

	/**
	 * Get file
	 * 
	 * @return
	 * @throws FileNotFoundException
	 */	
	public File getFile() throws FileNotFoundException {
		File here = new File(filepath);
		return new File(here.getAbsolutePath());
	}

	/**
	 * Get column names
	 * 
	 * @param workbook
	 * @param sheetIndex
	 * @return
	 */
	public ArrayList<String> getColNames(Workbook workbook, int sheetIndex) {
		ArrayList<String> colNames = new ArrayList<String>();
		try {
			Sheet sheet = workbook.getSheetAt(sheetIndex);
			Row row = sheet.getRow(0);
			int cols = 0;
			if (row != null) {
				cols = row.getPhysicalNumberOfCells();
				for (int i = 0; i < cols; i++) {
					colNames.add(getDataCell(workbook, row.getCell(i)));
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return colNames;
	}
	private List<String> getObjectFields(Object classname) {
		List<String> arr = new ArrayList<String>();
		Class cls = classname.getClass();
		Field[] fields = cls.getDeclaredFields();
		for (Field field : fields) {
			arr.add(field.getName());
		}
			
		return arr;
		
	}
	private boolean isEmptyRow(Row row){
		 boolean isEmptyRow = true;
		     for(int cellNum = row.getFirstCellNum(); cellNum < row.getLastCellNum(); cellNum++){
		        Cell cell = row.getCell(cellNum);
		        if(cell != null && cell.getCellType() != Cell.CELL_TYPE_BLANK ){
		        isEmptyRow = false;
		        }    
		     }
		 return isEmptyRow;
   }
	
	/**
	 * Mapping data with column name
	 * 
	 * @param workbook
	 * @param sheetIndex
	 * @return
	 */
	@SuppressWarnings("rawtypes")
	public ArrayList<Map> getMappedValues(Workbook workbook, int sheetIndex, Object modelName) {
		ArrayList<Map> mapArray = null;
		try {
			ArrayList<String> colNames = null;
			List<String> modelNames = getObjectFields(modelName);
			Row row = null;
			Sheet sheet = null;
			int sheetRows = 0;
			int rowCols = 0;
			Map<String, Object> rowMap = null;
			sheet = workbook.getSheetAt(sheetIndex);
			sheetRows = sheet.getPhysicalNumberOfRows();
			mapArray = new ArrayList<Map>(sheetRows - 1);
			colNames = getColNames(workbook, sheetIndex);
			colNames.trimToSize();
			rowCols = colNames.size();
			
			for (int i = 1; i < sheetRows; i++) {
				row = sheet.getRow(i);
				if(isEmptyRow(row)) {
					continue;
				}
				rowMap = new HashMap<String, Object>(rowCols);
				for (int c = 0; c < rowCols; c++) {
					String value = getDataCell(workbook, row.getCell(c));
				
					String colString = modelNames.get(c);
					
					rowMap.put(colString, value);
				}
				mapArray.add(rowMap);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return mapArray;
	}

	/**
	 * 
	 * @param workbook
	 * @param cell
	 * @return
	 */
	public String getDataCell(Workbook workbook, Cell cell) {
		String cellValue = null;
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				cellValue = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				cellValue = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				if (DateUtil.isCellInternalDateFormatted(cell)) {
					 Date date= cell.getDateCellValue();
					 cellValue = DtFormat.format(date).toString();
				} else {
					BigDecimal valueData = BigDecimal.valueOf(cell.getNumericCellValue());
					if(valueData != null){
						cellValue = valueData.toString();
					}
				}
				break;
			case Cell.CELL_TYPE_BLANK:
				cellValue = "";
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				cellValue = Boolean.toString(cell.getBooleanCellValue());
				break;
			}
		}
		return cellValue;
	}
	
	public ArrayList<String> getHeaderTemplate(String type) {
		ArrayList<String> colNames = null;
		try {
//			String pathTemplate = prop.getProperty("PATH_THEMPLATE");
//			String fileTemplateName = null;
//			// if upload adjustment
//			if(type.equals("ADJ")){
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_ADJ;
//			}
//			// if account mapping upload
//			else if (type.equals("ACCP")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_ACCP;
//			}
//			// if Agent Movement Other
//			else if (type.equals("AOMU")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_AOMU;
//			}
//			// if Upload Transfer Agent 
//			else if (type.equals("UTA")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_UTA;
//			}
//			// if Upload Demotion Agent 
//			else if (type.equals("UDA")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_UDA;
//			}
//			// if Upload Promotion Agent 
//			else if (type.equals("UPA")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_UPA;
//			}
//			// if Upload Termination Agent 
//			else if (type.equals("UTEA")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_UTEA;
//			}
//			// if Upload Transfer Policy
//			else if (type.equals("UTP")){ 
//				fileTemplateName = DCMSCode.UPLOAD_TEMPLATE_UTP;
//			}
//			String pathTemplateFile = pathTemplate + fileTemplateName;
//			File here = new File(pathTemplateFile);
//			FileInputStream file = new FileInputStream(here);
//			Workbook workbook = null;
//			if (fileType.equalsIgnoreCase("xlsx")) {
//				workbook = new XSSFWorkbook(file);
//			} else if (fileType.equalsIgnoreCase("xls")) {
//				workbook = new HSSFWorkbook(file);
//			}
//			int sheetIndex = 0;
//			colNames = getColNames(workbook, sheetIndex);
//			colNames.trimToSize();
//			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return colNames;
	}
	
	public ArrayList<String> getHeaderUpload() {
		ArrayList<String> colNames = null;
		try {
			FileInputStream file = new FileInputStream(getFile());
			Workbook workbook = null;
			if (fileType.equalsIgnoreCase("xlsx")) {
				workbook = new XSSFWorkbook(file);
			} else if (fileType.equalsIgnoreCase("xls")) {
				workbook = new HSSFWorkbook(file);
			}
			int sheetIndex = 0;
			colNames = getColNames(workbook, sheetIndex);
			colNames.trimToSize();
			file.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return colNames;
	}
	
	public Boolean isCorrectTemplate (String type){
		ArrayList<String> headerTemplateADJ = getHeaderTemplate(type);
		ArrayList<String> headerADJUpload = getHeaderUpload();
		Boolean isCorrect = true;
		if (headerADJUpload.size() != headerTemplateADJ.size()){
			isCorrect = false;
			return isCorrect;
		}
		for (int i = 0; i < headerTemplateADJ.size(); i++) {
			if(!headerADJUpload.get(i).equals(headerTemplateADJ.get(i))){
				isCorrect = false;
			}
		}
		return isCorrect;
	}

	
}
