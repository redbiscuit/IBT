package com.bj.utils.dataMgr;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.Iterator;


@SuppressWarnings({ "unchecked", "deprecation" })
public class ExcelFile {
	
	static String prjectPath = System.getProperty("user.dir")+"/";
	
	public static void main(String[] args){
		System.out.println(prjectPath);
		ExcelFile.writeExcel("datapool/GPE_Position.xls", "OEM", "OEMHardware_UAT", "Description", "555");
//		Data datapool = new Data();
//		datapool.setExeclInput("src/main/resources/datapool/GPE_MRT.xls");
//		System.out.println(ExcelFile.readMRT(datapool.getData("Currency", "Get_Currency_B")));
//		loopMrt("UAT-NV-64");
//		System.out.println(ExcelFile.getLatestVersion("UAT-NV-64", "Inflation"));
//		ExcelFile.writeExcel("datapool/MRT_Info.xls", "MRT_Info", 4, 5, "AAA");
//		ExcelFile.writeExcel("datapool/GPE_Position.xls", "Labor", "IBMBanded", "BCC","121","2");
//		System.out.println(ExcelFile.readExcel("src/main/resources/datapool/GPE_Summary.xls", "Summary", 0));
//		ExcelFile.writeExcel("src/main/resources/datapool/GPE_Summary.xls", "Summary", "IBMBanded", "TotalPrice", "420000");
		
//		ExcelFile.resetSheetFormula("src/main/resources/datapool/GPE_Summary.xls", "Summary");
//		ExcelFile.resetSheetFormula("datapool/GPE_Summary.xls", "ZPTI");
//		ExcelFile.resetSheetFormula("datapool/GPE_Summary.xls", "Summary");
		ExcelFile.resetSummarySheetFormula();
/*		Data Mrtpool = new Data();
		Mrtpool.setExeclInput("datapool/GPE_CMS_STD.xls");
		logutil.logInfo(ExcelFile.readMRT_2007(Mrtpool.getData("MRT", "Get_COLA_Inflation"), "INFL_YR2")) ;*/
	}
	
	public static void resetSummarySheetFormula(){
		resetSheetFormula("datapool/ST_汇总.xls", "Summary");
		resetSheetFormula("datapool/ST_汇总.xls", "History");
	}
	
	public static void resetSheetFormula(String fileName, String sheetName){
		fileName = prjectPath + fileName;
		File file = new File(fileName);
		FileInputStream in = null;
		String[] s = null;
		String temp = "";
		int l = 0;
		FormulaEvaluator eval=null; 
		
		try {
			in = new FileInputStream(file);
			HSSFWorkbook wb = new HSSFWorkbook(in);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFCell cell = null;
			
			if(wb instanceof HSSFWorkbook)  
	            eval=new HSSFFormulaEvaluator((HSSFWorkbook) wb);   
			
			for(int i = 0;i<sheet.getPhysicalNumberOfRows();i++){
				if(sheet.getRow(i)!= null){
					for(int j = 0;j<20;j++){
						cell = sheet.getRow(i).getCell(j);
						if(cell!=null&&cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
							 cell.setCellFormula(cell.getCellFormula());
						}
					}
				}
			}
			
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();
			
		} catch (Exception e) {
			System.out.println("" + file.getAbsolutePath() + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		
	}

	public static String[] readExcelData(String fileName, String sheetName, int rowNumber) {
		fileName = prjectPath + fileName;
		File file = new File(fileName);
		FileInputStream in = null;
		String[] s = null;
		String temp = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);

			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;
			int rowNum = rowNumber;

			row = sheet.getRow((short) rowNum);
			l = row.getLastCellNum();
			s = new String[l];

			for (int i = 0; i < l; i++) {
				cell = row.getCell((short) i);
				if (cell != null) {
					if (cell.getCellType() == 1)
						temp = cell.getStringCellValue();
					else
						temp = String.valueOf((int) cell.getNumericCellValue());
					s[i] = temp;
				} else {
					s[i] = "";
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("" + file.getAbsolutePath() + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return s;
	}

	public static String readExcel(String fileName, String sheetName, int rowNumber) {
		fileName = prjectPath + fileName;

		File file = new File(fileName);
		FileInputStream in = null;
		String s = "";
		String temp = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);

			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;
			int rowNum = rowNumber;

			row = sheet.getRow((short) rowNum);
			l = row.getLastCellNum();

			for (int i = 0; i < l; i++) {
				cell = row.getCell((short) i);
				if (cell != null) {
					if (cell.getCellType() == 1)
						temp = cell.getStringCellValue();
					else
						temp = String.valueOf((int) cell.getNumericCellValue());

					if (i == 0)
						s = temp;
					else
						s = s + "," + temp;
				} else {
					if (i == 0)
						s = "";
					else
						s = s + "," + " ";
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return s;
	}
	
	public static String readExcel_(String fileName, String sheetName, int rowNumber) {
//		fileName = prjectPath + fileName;

		File file = new File(fileName);
		FileInputStream in = null;
		String s = "";
		String temp = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);

			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;
			int rowNum = rowNumber;

			row = sheet.getRow((short) rowNum);
			l = row.getLastCellNum();

			for (int i = 0; i < l; i++) {
				cell = row.getCell((short) i);
				if (cell != null) {
					if (cell.getCellType() == 1)
						temp = cell.getStringCellValue();
					else
						temp = String.valueOf((int) cell.getNumericCellValue());

					if (i == 0)
						s = temp;
					else
						s = s + "," + temp;
				} else {
					if (i == 0)
						s = "";
					else
						s = s + "," + " ";
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return s;
	}

	public static HashMap readExcelWithHashMap(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;

		HashMap hm = new HashMap();
		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		String temp1 = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				row1 = sheet.getRow((short) (i + 1));
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							l = row.getLastCellNum();
							for (int ii = 1; ii < l; ii++) {
								cell = row.getCell((short) ii);
								cell1 = row1.getCell((short) ii);
								if (cell != null) {
									if (cell.getCellType() == 1)
										temp = cell.getStringCellValue();
									else
										temp = String.valueOf((int) cell.getNumericCellValue());

									if (!temp.trim().equals("")) {
										if (cell1 == null) {
											temp1 = " ";
										} else {
											if (cell1.getCellType() == 1)
												temp1 = cell1.getStringCellValue();
											else {
												temp1 = String.valueOf((int) cell1.getNumericCellValue());
												if (temp1.equals("0"))
													temp1 = " ";
											}
										}
										hm.put(temp, temp1);
									}
								}
							}
							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel Path " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return hm;
	}
	
	public static String loopMrt(String Platform) {
		String Path = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\";
		String MRT_Info = "datapool/MRT_Info.xls";
		String MRT_Sheet = "MRT_Info";
		String MRT_Sheet1 = "MRT_Test";
		
		File newfile = new File(prjectPath+MRT_Info);
		if(newfile.exists()){
     	   newfile.delete();
        }
		try{
		FileInputStream fs = null;
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet newSheet = null;
        newSheet = wb.createSheet(MRT_Sheet);
        HSSFSheet newSheet1 = null;
        newSheet1 = wb.createSheet(MRT_Sheet1);
        FileOutputStream out = new FileOutputStream(newfile);
        wb.write(out);
        out.close();
		}catch(Exception e){
			e.printStackTrace();
		}
		
		File file = new File(Path);
		File[] files = file.listFiles();
		ArrayList AL_Files = new ArrayList();
		String result;
		for(File f:files){
			if(f!=null && f.isFile()){
				String name = f.getName();
				if(name.endsWith(".xls")){
					if(name.indexOf("(")<0){
//						AL_Files.add(name.substring(0,name.indexOf(".xls")));
					}else{
						name = name.substring(0, name.indexOf("("));
						if(!AL_Files.contains(name)){
							AL_Files.add(name);
						}
					}
				}
			}
		}
		try{
		Iterator it = AL_Files.iterator();
		int i_row = 0;
		int i_col = 0;
		while(it.hasNext()){
			i_col = 0;
			String name = (String) it.next();
			String excelPath = Path + ExcelFile.getLatestVersion(Platform,name);
			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(excelPath));
			ExcelFile.writeExcel(MRT_Info, MRT_Sheet, i_row, i_col, ExcelFile.getLatestVersion(Platform,name));
//			System.out.println(ExcelFile.getLatestVersion(Platform,name));
			for (int numSheets = 0; numSheets < workbook.getNumberOfSheets(); numSheets++) {
				if (null != workbook.getSheetAt(numSheets)) {
					HSSFSheet aSheet = workbook.getSheetAt(numSheets);
					int sheetLastRowNum = aSheet.getLastRowNum();
					String sheetName = aSheet.getSheetName();
					i_col++;
					ExcelFile.writeExcel(MRT_Info, MRT_Sheet, i_row, i_col, sheetName);
					
					ExcelFile.writeExcel(MRT_Info, MRT_Sheet, i_row+1, i_col, String.valueOf(sheetLastRowNum));
					
					/*	for (int rowNumOfSheet = 1; rowNumOfSheet <= aSheet.getLastRowNum(); rowNumOfSheet++){
						if (null != aSheet.getRow(rowNumOfSheet)) {
							HSSFRow aRow = aSheet.getRow(rowNumOfSheet);
						}
					}*/
				}
			}
			i_row++;
			i_row++;
		}
		}catch(Exception e){
			e.printStackTrace();
		}
		
		return "";
	}
	
	public static String getLatestVersion(String Platform,String MRTFile){
		String Path = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\";
		File file = new File(Path);
		File[] files = file.listFiles();
		ArrayList AL_Files = new ArrayList();
		String result = null;
		int temp = 1;
		
		for(File f:files){
			if(f!=null && f.isFile()){
				String name = f.getName();
				if((name.endsWith(".xls")||(name.endsWith(".xlsx")))&&name.startsWith(MRTFile)){
					AL_Files.add(name);
				}
			}
		}
		
		if(AL_Files.size()==1){
			result = AL_Files.get(0).toString();
		}else if(AL_Files.size()>1){
			Iterator it = AL_Files.iterator();
			while(it.hasNext()){
				String cur_version = it.next().toString();
				cur_version = cur_version.substring(cur_version.indexOf("(")+1, cur_version.indexOf(")"));
				if(Integer.parseInt(cur_version)>temp){
					temp = Integer.parseInt(cur_version);
				}
			}
			if(MRTFile.equals("CMS.ITSRates")){
				result = MRTFile + "(" + String.valueOf(temp) + ").xlsx";
			}else{
				result = MRTFile + "(" + String.valueOf(temp) + ").xls";
			}
		}else{
			return "";
		}
		
		return result;
	}
	
	public static String readMRT_2007(ArrayList<Hashtable<String,String>> al,String get_cell){
		Hashtable<String,String> ht;
		int con_len = al.size();
		ht = al.get(0);
		String Platform = ht.get("Platform");
		String MRTFile = ht.get("MRTFile");
		String SheetName = ht.get("Sheet");
		String GetCell = get_cell;
		String fileName = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\" + getLatestVersion(Platform,MRTFile);
		File file = new File(fileName);

		FileInputStream in = null;

		String result = "";
		int l = 0;
		
		ArrayList <Integer> Al_col_Num  = new ArrayList<Integer>();
		int get_col_true = 0;
		int get_row_true = 0;
		
		try{
//			in = new FileInputStream(file);
			OPCPackage opcPackage = OPCPackage.open(fileName);
			XSSFWorkbook workbook = new XSSFWorkbook(opcPackage);
//			XSSFWorkbook workbook = new XSSFWorkbook(in);
//			Workbook workBook1 = new SXSSFWorkbook(100);
			XSSFSheet sheet = workbook.getSheet(SheetName);
			
			XSSFRow row = null;
			XSSFRow row1 = null;
			XSSFRow row_R = null;
			XSSFCell cell = null;
			XSSFCell cell1 = null;
			XSSFCell cell_R = null;
		
			row = sheet.getRow(0);
			l = row.getLastCellNum();
			
			for(int i = 0; i < l; i++){
				cell = row.getCell(i);
				if(cell != null){
					if(cell.getStringCellValue().equals(GetCell)){
						get_col_true = i;
					}
					for(int j = 0; j < con_len; j++){
						if(cell.getStringCellValue().trim().equals(al.get(j).get("ColumnName").trim())){
							Al_col_Num.add(i);
						}
					}
				}
			}                
			
			for(int cur_row = 1; cur_row < sheet.getLastRowNum(); cur_row++){
				row1 = sheet.getRow(cur_row);
				cell1 = row1.getCell(Al_col_Num.get(0));
				
				int row_cal = 0;
				
				for(int col_num = 0;col_num<con_len;col_num++){
					cell1 = row1.getCell(Al_col_Num.get((Integer)col_num));
//					System.out.println("MRTValue:" + cell1.getStringCellValue() + " ExcelValue:" +al.get(col_num).get("ColumnValue"));
					if(cell1.getStringCellValue().trim().equals(al.get(col_num).get("ColumnValue").trim())){
						row_cal ++;
					}
				}
				
				if(row_cal == con_len){
					get_row_true = cur_row;
				}
			}
			
			if(get_row_true != 0){
				row_R = sheet.getRow(get_row_true);
				cell_R = row_R.getCell(get_col_true);
				result = cell_R.getStringCellValue();
			}

		}catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}
	
	public static String readMRT(ArrayList<Hashtable<String,String>> al,String get_cell){
		Hashtable<String,String> ht;
		int con_len = al.size();
		ht = al.get(0);
		String Platform = ht.get("Platform");
		String MRTFile = ht.get("MRTFile");
		String SheetName = ht.get("Sheet");
		String GetCell = get_cell;
		String fileName = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\" + getLatestVersion(Platform,MRTFile);
		File file = new File(fileName);

		FileInputStream in = null;

		String result = "";
		int l = 0;
		
		ArrayList <Integer> Al_col_Num  = new ArrayList<Integer>();
		int get_col_true = 0;
		int get_row_true = 0;
		
		try{
			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(SheetName);
			
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFRow row_R = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;
			HSSFCell cell_R = null;
			
			row = sheet.getRow(0);
			l = row.getLastCellNum();
			
			for(int i = 0; i < l; i++){
				cell = row.getCell(i);
				if(cell != null){
					if(cell.getStringCellValue().equals(GetCell)){
						get_col_true = i;
					}
					for(int j = 0; j < con_len; j++){
						if(cell.getStringCellValue().trim().equals(al.get(j).get("ColumnName").trim())){
							Al_col_Num.add(i);
						}
					}
				}
			}                
			
			for(int cur_row = 1; cur_row < sheet.getLastRowNum(); cur_row++){
				row1 = sheet.getRow(cur_row);
				cell1 = row1.getCell(Al_col_Num.get(0));
				
				int row_cal = 0;
				
				for(int col_num = 0;col_num<con_len;col_num++){
					cell1 = row1.getCell(Al_col_Num.get((Integer)col_num));
//					System.out.println("MRTValue:" + cell1.getStringCellValue() + " ExcelValue:" +al.get(col_num).get("ColumnValue"));
					if(cell1.getStringCellValue().trim().equals(al.get(col_num).get("ColumnValue").trim())){
						row_cal ++;
					}
				}
				
				if(row_cal == con_len){
					get_row_true = cur_row;
				}
			}
			
			if(get_row_true != 0){
				row_R = sheet.getRow(get_row_true);
				cell_R = row_R.getCell(get_col_true);
				result = cell_R.getStringCellValue();
			}

		}catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}

	public static String readMRT(ArrayList<Hashtable<String,String>> al){
		Hashtable<String,String> ht;
		int con_len = al.size();
		ht = al.get(0);
		String Platform = ht.get("Platform");
		String MRTFile = ht.get("MRTFile");
		String SheetName = ht.get("Sheet");
		String GetCell = ht.get("GetCell");
		String fileName = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\" + getLatestVersion(Platform,MRTFile);
		File file = new File(fileName);
		FileInputStream in = null;

		String result = "";
		int l = 0;
		
		ArrayList <Integer> Al_col_Num  = new ArrayList<Integer>();
		int get_col_true = 0;
		int get_row_true = 0;
		
		try{
			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(SheetName);
			
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFRow row_R = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;
			HSSFCell cell_R = null;
			
			row = sheet.getRow(0);
			l = row.getLastCellNum();
			
			for(int i = 0; i < l; i++){
				cell = row.getCell(i);
				if(cell != null){
					if(cell.getStringCellValue().equals(GetCell)){
						get_col_true = i;
					}
					for(int j = 0; j < con_len; j++){
						if(cell.getStringCellValue().trim().equals(al.get(j).get("ColumnName").trim())){
							Al_col_Num.add(i);
						}
					}
				}
			}                
			
			for(int cur_row = 1; cur_row < sheet.getLastRowNum(); cur_row++){
				row1 = sheet.getRow(cur_row);
				cell1 = row1.getCell(Al_col_Num.get(0));
				
				int row_cal = 0;
				
				for(int col_num = 0;col_num<con_len;col_num++){
					cell1 = row1.getCell(Al_col_Num.get((Integer)col_num));
//					System.out.println("MRTValue:" + cell1.getStringCellValue() + " ExcelValue:" +al.get(col_num).get("ColumnValue"));
					if(cell1.getStringCellValue().trim().equals(al.get(col_num).get("ColumnValue").trim())){
						row_cal ++;
					}
				}
				
				if(row_cal == con_len){
					get_row_true = cur_row;
				}
			}
			
			if(get_row_true != 0){
				row_R = sheet.getRow(get_row_true);
				cell_R = row_R.getCell(get_col_true);
				result = cell_R.getStringCellValue();
			}

		}catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}
	
	public static String readMRT(String Platform,ArrayList<Hashtable<String,String>> al){
		Hashtable<String,String> ht;
		int con_len = al.size();
		ht = al.get(0);
		String MRTFile = ht.get("MRTFile");
		String SheetName = ht.get("Sheet");
		String GetCell = ht.get("GetCell");
		String fileName = "C:\\SDM\\GPE\\"+Platform+"\\workspace\\" + getLatestVersion(Platform,MRTFile);
		File file = new File(fileName);
		FileInputStream in = null;

		String result = "";
		int l = 0;
		
		ArrayList <Integer> Al_col_Num  = new ArrayList<Integer>();
		int get_col_true = 0;
		int get_row_true = 0;
		
		try{
			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(SheetName);
			
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFRow row_R = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;
			HSSFCell cell_R = null;
			
			row = sheet.getRow(0);
			l = row.getLastCellNum();
			
			for(int i = 0; i < l; i++){
				cell = row.getCell(i);
				if(cell != null){
					if(cell.getStringCellValue().equals(GetCell)){
						get_col_true = i;
					}
					for(int j = 0; j < con_len; j++){
						if(cell.getStringCellValue().trim().equals(al.get(j).get("ColumnName").trim())){
							Al_col_Num.add(i);
						}
					}
				}
			}                
			
			for(int cur_row = 1; cur_row < sheet.getLastRowNum(); cur_row++){
				row1 = sheet.getRow(cur_row);
				cell1 = row1.getCell(Al_col_Num.get(0));
				
				int row_cal = 0;
				
				for(int col_num = 0;col_num<con_len;col_num++){
					cell1 = row1.getCell(Al_col_Num.get((Integer)col_num));
//					System.out.println("MRTValue:" + cell1.getStringCellValue() + " ExcelValue:" +al.get(col_num).get("ColumnValue"));
					if(cell1.getStringCellValue().trim().equals(al.get(col_num).get("ColumnValue").trim())){
						row_cal ++;
					}
				}
				
				if(row_cal == con_len){
					get_row_true = cur_row;
				}
			}
			
			if(get_row_true != 0){
				row_R = sheet.getRow(get_row_true);
				cell_R = row_R.getCell(get_col_true);
				result = cell_R.getStringCellValue();
			}

		}catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}
	
	public static ArrayList readExcelReturnArrayListIncludeDataKey(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;
		ArrayList hm = new ArrayList();
		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		String temp1 = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				row1 = sheet.getRow((short) (i + 1));
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							String[] t = new String[2];
							t[0] = title;
							t[1] = "";
							hm.add(t);

							l = row.getLastCellNum();
							for (int ii = 1; ii < l; ii++) {
								cell = row.getCell((short) ii);
								cell1 = row1.getCell((short) ii);
								if (cell != null) {
									if (cell.getCellType() == 1)
										temp = cell.getStringCellValue();
									else
										temp = String.valueOf((int) cell.getNumericCellValue());

									if ((!temp.trim().equals("")) && (!temp.trim().equals("0"))) {
										if (cell1 == null) {
											temp1 = " ";
										} else {
											if (cell1.getCellType() == 1)
												temp1 = cell1.getStringCellValue();
											else {
												temp1 = String.valueOf((int) cell1.getNumericCellValue());
												if (temp1.equals("0"))
													temp1 = " ";
											}
										}
										String[] sa = new String[2];
										sa[0] = temp;
										sa[1] = temp1;
										hm.add(sa);
									}
								}
							}
							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return hm;
	}

	public static ArrayList readExcelReturnArrayList(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;
		ArrayList hm = new ArrayList();
		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		String temp1 = "";
		int l = 0;
		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				row1 = sheet.getRow((short) (i + 2));
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							l = row.getLastCellNum();
							for (int ii = 1; ii < l; ii++) // control ---------
							{
								cell = row.getCell((short) ii);
								cell1 = row1.getCell((short) ii);
								if (cell != null) {
									if (cell.getCellType() == 1)
										temp = cell.getStringCellValue();
									else
										temp = String.valueOf((int) cell.getNumericCellValue());

									if ((!temp.trim().equals("")) && (!temp.trim().equals("0"))) {
										if (cell1 == null) {
											temp1 = " ";
										} else {
											if (cell1.getCellType() == 1)
												temp1 = cell1.getStringCellValue();
											else {
												temp1 = String.valueOf((int) cell1.getNumericCellValue());
												if (temp1.equals("0"))
													temp1 = " ";
											}
										}
										String[] sa = new String[2];
										sa[0] = temp;
										sa[1] = temp1;
										// System.out.println(temp);
										// System.out.println(temp1);
										hm.add(sa);
									}
								}
							}
							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return hm;
	}

	public static ArrayList readExcelWithStringArray(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;
		File file = new File(fileName);
		FileInputStream in = null;
		ArrayList list = new ArrayList();

		String temp = "";
		int l = 0;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							l = row.getLastCellNum();
							for (int ii = 1; ii < l; ii++) {
								cell = row.getCell((short) ii);
								if (cell != null) {
									if (cell.getCellType() == 1)
										temp = cell.getStringCellValue();
									else {
										temp = String.valueOf((int) cell.getNumericCellValue());
										if (temp.equals("0"))
											temp = "";
									}

									if (!temp.trim().equals(""))
										list.add(temp);

								}
							}

							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println(file.getAbsolutePath() + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return list;
	}

	public static String readExcel(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;
		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		String result = "";

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							cell = row.getCell((short) 3);
							if (cell.getCellType() == 1)
								result = cell.getStringCellValue();
							else
								result = String.valueOf((int) cell.getNumericCellValue());

							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}

	public static short readExcel(String fileName, String sheetName, String title, short cellNumber) {
		fileName = prjectPath + fileName;
		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		short result = -1;

		try {

			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell(cellNumber);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title.trim())) {
							result = (short) i;
							break;
						}
					}
				}
			}

			in.close();
		} catch (Exception e) {
			System.out.println(file.getAbsolutePath() + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return result;
	}


	public static void writeExcel(String fileName, String sheetName, int rowNumber, int columnNumber, String s) {

		FileOutputStream fOut = null;
		try {

			fileName = prjectPath + fileName;

			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(fileName));

			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFRow row = sheet.getRow(rowNumber);
			if (row == null)
				row = sheet.createRow(rowNumber);

			HSSFCell cell = row.getCell((short) columnNumber);
			if (cell == null)
				cell = row.createCell((short) columnNumber);

			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			// cell.setCellType(HSSFCell.ENCODING_UTF_16);
			cell.setCellValue(s);
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			System.out.println("Excel File Write Error: " + e);
		} finally {
			if (fOut != null) {
				try {
					fOut.close();
				} catch (IOException e1) {
				}
			}
		}
	}

	public static void writeExcel(String fileName, String sheetName,String title, String value) {
		FileOutputStream fOut = null;
		String temp;
		try {

			fileName = prjectPath + fileName;
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					fileName));

			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							row1 = sheet.getRow((short) (i + 1));
							if (row1 == null)
								row1 = sheet.createRow((short) (i + 1));
							cell = row1.getCell((short) 1);
							if (cell == null) {
								cell = row1.createCell((short) 1);
							}
							cell.setCellType(HSSFCell.CELL_TYPE_STRING);
							cell.setCellValue(value);

							break;
						}
					}
				}
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			System.out.println("Excel File Write Error: " + e);
		} finally {
			if (fOut != null) {
				try {
					fOut.close();
				} catch (IOException e1) {
				}
			}
		}
	}

	public static void writeExcel(String fileName, String sheetName,String title, String key, String value) {

		FileOutputStream fOut = null;
		String temp;
		try {
			fileName = prjectPath + fileName;
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					fileName));

			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							short l = row.getLastCellNum();
							short r = (short) (i + 1);
							short c = 0;
							for (int n = 1; n < l; n++) {
								cell = row.getCell((short) n);
								if (cell != null) {
									temp = cell.getStringCellValue();
									if (temp.trim().equals(key.trim())) {
										c = (short) n;
										break;
									}
								}
							}
							if (c != 0) {
								row1 = sheet.getRow((short) r);
								if (row1 == null)
									row1 = sheet.createRow((short) r);
								cell = row1.getCell((short) c);
								if (cell == null) {
									cell = row1.createCell((short) c);
								}
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								cell.setCellValue(value);
								break;
							}
						}
					}
				}
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();
//			System.out.println("Excel File Write success!" );

		} catch (Exception e) {
			System.out.println("Excel File Write Error: " + e);
		} finally {
			if (fOut != null) {
				try {
					fOut.close();
				} catch (IOException e1) {
				}
			}
		}
	}

	public static void writeExcel(String fileName, String sheetName,String title, String key, String value,String rowIndex) {

		FileOutputStream fOut = null;
		String temp;
		try {
			fileName = prjectPath + fileName;
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					fileName));

			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;

			for (int i = 0; i < 1000; i++) {
				row = sheet.getRow((short) i);
				if (row != null) {
					cell = row.getCell((short) 0);
					if (cell != null) {
						temp = cell.getStringCellValue();
						if (temp.trim().equals(title)) {
							short l = row.getLastCellNum();
							short r = (short) (i + Short.parseShort(rowIndex));
							short c = 0;
							for (int n = 1; n < l; n++) {
								cell = row.getCell((short) n);
								if (cell != null) {
									temp = cell.getStringCellValue();
									if (temp.trim().equals(key.trim())) {
										c = (short) n;
										break;
									}
								}
							}
							if (c != 0) {
								row1 = sheet.getRow((short) r);
								if (row1 == null)
									row1 = sheet.createRow((short) r);
								cell = row1.getCell((short) c);
								if (cell == null) {
									cell = row1.createCell((short) c);
								}
								cell.setCellType(HSSFCell.CELL_TYPE_STRING);
								cell.setCellValue(value);
								break;
							}
						}
					}
				}
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			System.out.println("Excel File Write Error: " + e);
		} finally {
			if (fOut != null) {
				try {
					fOut.close();
				} catch (IOException e1) {
				}
			}
		}
	}
	
	public static void writeExcel(String fileName, String sheetName,int rowNumber, String s) {

		FileOutputStream fOut = null;
		try {

			fileName = prjectPath + fileName;

			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					fileName));

			String[] ss = s.split(",");
			int l = ss.length;

			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheet(sheetName);
			HSSFRow row = sheet.getRow(rowNumber);
			if (row == null)
				row = sheet.createRow(rowNumber);

			HSSFCell cell = null;
			for (int i = 0; i < l; i++) {
				cell = row.getCell((short) i);
				if (cell == null)
					cell = row.createCell((short) i);

				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				// cell.setCellType(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(ss[i]);
			}
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(fileName);
			wb.write(fileOut);
			fileOut.close();

		} catch (Exception e) {
			System.out.println("Excel File Write Error: " + e);
		} finally {
			if (fOut != null) {
				try {
					fOut.close();
				} catch (IOException e1) {
				}
			}
		}
	}

//	public static void copyBaseFile(String newFileName) {
//		final String sfilename = Config.getProperty("PROJECT_PATH") + "\\testreports\\TestReport_base.xls";
//
//		try {
//			Runtime rt = Runtime.getRuntime();
//			Process proc = rt.exec("cmd /c copy \"" + sfilename + "\" \"" + newFileName + "\"");
//			int exitVal = proc.waitFor();
//		} catch (Exception e) {
//			// TODO: handle exception
//		}
//	}

	public static ArrayList readExcelReturnArrayList4MultiData(String fileName, String sheetName, String title) {
		fileName = prjectPath + fileName;
		ArrayList hm = new ArrayList();

		ArrayList rhm = new ArrayList();
		ArrayList rhm2 = null;

		ArrayList a = new ArrayList();

		File file = new File(fileName);
		FileInputStream in = null;
		String temp = "";
		String temp1 = "";
		int l = 0;
		try {
			/*
			 * 有了HSSFWorkbook实例，接下来就可以提取工作表、工作表的行和列，例如：
			 * 
			 * HSSFSheet sheet = wb.getSheetAt(0); // 第一个工作表 HSSFRow row = sheet.getRow(2); // 第三行 HSSFCell cell =
			 * row.getCell((short)3); // 第四个单元格
			 */
			in = new FileInputStream(file);
			HSSFWorkbook workbook = new HSSFWorkbook(in);
			HSSFSheet sheet = workbook.getSheet(sheetName);

			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;

			for (int i = 0; i < 1000; i++) {
				int colSize = sheet.getLastRowNum();// 列数，从0开始编号
				for (int j = 1; j < colSize + 1; j++) {
					row = sheet.getRow((short) i);
					row1 = sheet.getRow((short) (i + j));// 获取实际数据列表，需要判断是否存在数据

					if (row != null && row1 != null) {
						cell = row.getCell((short) 0);
						if (cell != null) {
							temp = cell.getStringCellValue();
							if (temp.trim().equals(title)) // 不等于title就直接中断
							{
								l = row.getLastCellNum();
								for (int ii = 1; ii < l; ii++) // control
								// ---------行实际数量
								{
									cell = row.getCell((short) ii);
									cell1 = row1.getCell((short) ii);
									if (cell != null) {
										if (cell.getCellType() == 1) {
											temp = cell.getStringCellValue();
										} else {
											temp = String.valueOf((int) cell.getNumericCellValue());
										}
										if ((!temp.trim().equals("")) && (!temp.trim().equals("0"))) {
											if (cell1 == null) {
												temp1 = " ";
											} else {
												if (cell1.getCellType() == HSSFCell.CELL_TYPE_STRING) {
													temp1 = cell1.getStringCellValue();
											//		System.out.println("AAA"+temp1);
												}
												// added for float number
												// process
												else if (cell1.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
													temp1 = String.valueOf(cell1.getNumericCellValue());
													
												} else {
													temp1 = String.valueOf((int) cell1.getNumericCellValue());
											//		System.out.println("BBB"+temp1);
													if (temp1.equals("0")) {
														temp1 = " ";
													}

												}
											}
											String[] sa = new String[2];
											sa[0] = temp;
											sa[1] = temp1;
						//					 System.out.println("AAA1"+temp);
						//					 System.out.println("BBB1"+temp1);
											hm.add(sa);

										}
									}

								}
								rhm.add(hm.clone());// 不用clone则引用被清空全清空
								hm.clear();// vacant temp container
							}
						}
					} else {
						break;
					}

				}

			}

			in.close();
		} catch (Exception e) {
			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return rhm;
	}

	public static ArrayList readExcelReturnArrayList4MultiDataTest(String fileName, String sheetName, String title) {

		ArrayList hm = new ArrayList();

		ArrayList rhm = new ArrayList();
		ArrayList rhm2 = null;

		ArrayList a = new ArrayList();

		File file = new File(fileName);
		
		FileInputStream in = null;
		String temp = "";
		String temp1 = "";
		int l = 0;
		try {
			/*
			 * 有了HSSFWorkbook实例，接下来就可以提取工作表、工作表的行和列，例如：
			 * 
			 * HSSFSheet sheet = wb.getSheetAt(0); // 第一个工作表 HSSFRow row = sheet.getRow(2); // 第三行 HSSFCell cell =
			 * row.getCell((short)3); // 第四个单元格
			 */
			in = new FileInputStream(file);

			HSSFWorkbook workbook = new HSSFWorkbook(in);
			
			HSSFSheet sheet = workbook.getSheet(sheetName);
			
			HSSFFormulaEvaluator e = new HSSFFormulaEvaluator(workbook);

			HSSFRow row = null;
			HSSFRow row1 = null;
			HSSFCell cell = null;
			HSSFCell cell1 = null;
			
	//		logutil.logInfo("Total Rows in Summary Sheet: "+sheet.getLastRowNum());
			LOOP:
			for (int i = 0; i < 400; i++) {
				int colSize = sheet.getLastRowNum();// 列数，从0开始编号

				for (int j = 1; j < colSize + 1; j++) {
					row = sheet.getRow((short) i);
					row1 = sheet.getRow((short) (i + j));// 获取实际数据列表，需要判断是否存在数据

					if (row == null && row1 == null) {
						break;
					}

					if (row != null && row1 != null) {
						cell = row.getCell((short) 0);
						if (cell != null) // 如果存在title
						{
				//			System.out.println("Title is exist");
							if (cell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
								temp = cell.getStringCellValue();
							}else if (cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
								cell = e.evaluateInCell(cell);
								temp = String.valueOf(cell.getStringCellValue());
							}
				//			System.out.println(j+temp);
							if (temp.trim().equals(title)) // 不等于title就直接中断
							{
								l = row.getLastCellNum();
//								System.out.println(l);
								for (int ii = 1; ii < l; ii++) // control
								// ---------行实际数量
								{
						//			System.out.println(ii);
									cell = row.getCell((short) ii);
									cell1 = row1.getCell((short) ii);

									
				//					if (cell != null && cell1 != null && !cell1.getStringCellValue().equals(""))// 只要实际数据开头一项有空值则退出
									if(cell1 == null){
//										System.out.println(ii);
//									System.out.println(cell1.getStringCellValue());
								}

									if (cell != null && cell1 != null )
									{
										if (cell.getCellType() == 1) {
//											System.out.println(temp);
											temp = cell.getStringCellValue();
										} else if(cell.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
											cell = e.evaluateInCell(cell);
											temp = String.valueOf( cell.getStringCellValue());
//											System.out.println(temp);
										}

										if ((!temp.trim().equals("")) && (!temp.trim().equals("0"))) {

											if (cell1 == null) {
												temp1 = " ";
											} else {
												if (cell1.getCellType() == HSSFCell.CELL_TYPE_STRING) {
													temp1 = cell1.getStringCellValue();
													if(temp1.equals("END")){
														break LOOP;
													}
//													System.out.println("OOO"+temp1);
												}
												// added for float number
												// process
												else if (cell1.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
													temp1 = String.valueOf(cell1.getNumericCellValue());
									//				System.out.println("AAA"+temp1);
												} else if(cell1.getCellType() == HSSFCell.CELL_TYPE_FORMULA){
													cell1 = e.evaluateInCell(cell1);
													if(cell1.getCellType() == 0){
														temp1 = String.valueOf(cell1.getNumericCellValue());}
													else if(cell1.getCellType() == 1){
														temp1 = String.valueOf(cell1.getStringCellValue());}
													
												} else {
													temp1 = String.valueOf((Double) cell1.getNumericCellValue());
									//				System.out.println("BBB"+temp1);
													if (temp1.equals("0")) {
														temp1 = " ";
													}

												}
											}
											String[] sa = new String[2];
											sa[0] = temp;
											sa[1] = temp1;
											hm.add(sa);
										}
									}

								}
								if (hm.size() > 0) {
									rhm.add(hm.clone());// 不用clone则引用被清空全清空
									hm.clear();// vacant temp container
								} else// 如果遇到空行，则认为此title已经处理结束
								{
									break;
								}
							}
						}
					} else {
						break;
					}

				}

			}

			in.close();
		} catch (Exception e) {
//			System.out.println("Excel File " + file.getAbsolutePath() + " read error: " + e);
			e.printStackTrace();
		} finally {
			if (in != null) {
				try {
					in.close();
				} catch (IOException e1) {
				}
			}
		}
		return rhm;
	}

}

