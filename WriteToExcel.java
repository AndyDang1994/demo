package poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.codehaus.jackson.map.ObjectMapper;

public class WriteToExcel {
	
	private String filePath;
	private String fileType;
	private Sheet sheet;
	private Workbook workbook;
	
	public WriteToExcel(String filePath) {
		
		if(filePath != null) {
			if (filePath.endsWith("xlsx")) {
				fileType = "xlsx";
			} else if (filePath.endsWith("xls")) {
				fileType = "xls";
			}
		}
		
		this.filePath = filePath;
	}
	
	
	public WriteToExcel writeDataToExcel() {
		Workbook workbook = null;
		
        if (this.fileType.equals("xlsx")) {
            workbook = new XSSFWorkbook();
        } else if (this.fileType.equals("xls")) {
            workbook = new HSSFWorkbook();
        } else {
            throw new IllegalArgumentException("The specified file is not Excel file");
        } 
        this.sheet = workbook.createSheet("sheet1");
        this.workbook = workbook;
        return this;
        
	}
	
	public WriteToExcel writeHeader(List<String> headers) {
		
		Row row = this.sheet.createRow(0);
		int index = 0;
		for (String header : headers) {
			Cell cell = row.createCell(index);
			cell.setCellValue(header);
			cell.getCellStyle().setFillForegroundColor(IndexedColors.GREEN.getIndex());
			index++;
		}
		return this;
	}

	public WriteToExcel writeData( List<?> models) {
		int rowInd = 1;
		for (Object object : models) {
			ObjectMapper objectMapper = new ObjectMapper();
			Map<String, Object> convertValue = objectMapper.convertValue(object, Map.class);
			Set set = convertValue.entrySet();
			Iterator itr = set.iterator(); 
			int colindex = 0;
			Row row = this.sheet.createRow(rowInd);
			while (itr.hasNext()) {
				Map.Entry entry = (Map.Entry)itr.next();
				Object val = entry.getValue();
				Cell cell = row.createCell(colindex);
				
				if(val instanceof String) {
					cell.setCellValue((String)val);
				}
				
				colindex++;
				
			}
			rowInd++;
		}
		return this;
	}
	
	public void write() {
		
		try {
			OutputStream oStream = new FileOutputStream(this.filePath);
			this.workbook.write(oStream);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}
	
}
