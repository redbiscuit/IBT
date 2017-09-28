package com.bj.utils.dataMgr;

import java.util.ArrayList;
import java.util.Hashtable;

public class Data {

	/**
	 * @Description
	 * @author jiang.bian
	 * @param
	 * @return boolean
	 * @date Sep 11, 2013
	 */
	String ExeclInput;
	String ExcelVerify;


	public static void main(String[] args) {
		Data datapool = new Data();
		datapool.setExeclInput("datapool/ST_1.xls");
		ArrayList<Hashtable<String,String>> al;

		al = datapool.getSummaryData("分批赔率二");
		Hashtable<String,String> ht;

//		System.out.println(al.size());

		ht = al.get(0);
		System.out.println(ht.get("赔率上限"));

		// TODO Auto-generated method stub
//		ExcelFile.writeExcel("E:\\IBM_ADMIN\\IBM\\rationalsdp\\workspace\\GPE Regression Automation\\datapool\\GPE_Position.xls", "OEM", "OEMHardware_UAT", "Description", "555");
	}
	
	public void setExeclInput(String xls){
		this.ExeclInput = xls;
	}
	
	public String getExeclInput(){
		return this.ExeclInput;
	}
	
	public void setExcelVerify(String xls){
		this.ExcelVerify = xls;
	}
	
	public String getExcelVerify(){
		return this.ExcelVerify;
	}

	public ArrayList<Hashtable<String,String>> getSummaryData(String tablename){
	    ArrayList<Hashtable<String, String>> al;
		al = new ArrayList<Hashtable<String, String>>();		
		try {
			al = GetData.getGroupData(this.ExeclInput, "赔率设置", tablename);
	//		ht = al.get(0);
			if (al.size() > 0) {

			} else {
				System.out.println("no Data found from excel!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return al;
	}

	public ArrayList<Hashtable<String,String>> getData(String sheetname,String tablename){
	    ArrayList<Hashtable<String, String>> al;
		al = new ArrayList<Hashtable<String, String>>();		
		try {
			al = GetData.getGroupData(this.ExeclInput, sheetname, tablename);
	//		ht = al.get(0);
			if (al.size() > 0) {

			} else {
				System.out.println("no Data found from excel!");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return al;
	}
	


}
