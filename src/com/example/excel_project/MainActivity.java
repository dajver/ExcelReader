package com.example.excel_project;

import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import android.os.Bundle;
import android.app.Activity;
import android.util.Log;
import android.view.Menu;

public class MainActivity extends Activity {

	@Override
	protected void onCreate(Bundle savedInstanceState) {
		super.onCreate(savedInstanceState);
		setContentView(R.layout.activity_main);
		
		try {
			//��������� ����
			Workbook wb = WorkbookFactory.create(getAssets().open("test.xls"));
			//����� ������ ��������
         	Sheet sheet = wb.getSheetAt(0);
         	//��������� ���� ��� ������ �������������� �����, getPhysicalNumberOfRows() ����� ������ ����������
         	for(int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
         		Row row = sheet.getRow(i);
	         	//������ �������
	         	Cell name = row.getCell(0);
	         	Cell age = row.getCell(1);
	         	//���������� ��������� ������ + ��������
	         	Log.d("",name.getStringCellValue() + " " + age.getNumericCellValue());
         	}
         	
		 } catch(Exception ex) {
			 return;
		 }
	}
}
