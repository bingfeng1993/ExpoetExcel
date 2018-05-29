package com.dao.chu.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

public class TestImportExcel {

	public static void main(String[] args) throws IOException, Exception {

		String fileName = "student.xlsx";
		InputStream in = new FileInputStream(new File("excelfile\\student.xlsx"));
		Workbook wb = ImportExeclUtil.chooseWorkbook(fileName, in);
		StudentStatistics studentStatistics = new StudentStatistics();

		// 读取一个对象的信息
		StudentStatistics readDateT = ImportExeclUtil.readDateT(wb, studentStatistics, in, new Integer[] { 12, 5 },
				new Integer[] { 13, 5 });
		System.out.println(readDateT);

		// 读取对象列表的信息
		StudentBaseInfo studentBaseInfo = new StudentBaseInfo();
		// 第二行开始，到倒数第三行结束（总数减去两行）
		List<StudentBaseInfo> readDateListT = ImportExeclUtil.readDateListT(wb, studentBaseInfo, 2, 2);
		System.out.println(readDateListT);

	}
}
