package com.poi.xwpf;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map.Entry;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.junit.Test;

/**
 * 收集的官方代码示例：https://www.programcreek.com/java-api-examples/index.php?api=org.apache.poi.hwpf.HWPFDocument
 * 	看起来很厉害但是好像不怎么实用
 */
public class POIXwpfTest {

	/**
	 * 将word中的占位符${aaa}替换为aaa
	 */
	@Test
	public void wordReplaceTest() {
		HashMap<String, String> param = new HashMap<String, String>(); 
		param.put("a", "啊");
		param.put("aa", "啊啊");
		param.put("aaa", "啊啊啊");
		
		String inputPath = "C:/Users/30616/Desktop/temp/test.doc";
		String outputPath = "C:/Users/30616/Desktop/temp/testResult.doc";
		InputStream is = null;
		OutputStream os = null;
		
		try {
			is = new FileInputStream(inputPath);
			HWPFDocument doc = new HWPFDocument(is);
			Range range = doc.getRange();
			
			for (Entry<String, String> entry : param.entrySet()) {
				range.replaceText("${"+entry.getKey().toLowerCase()+"}", entry.getValue());
			}
			
			os = new FileOutputStream(outputPath);
			doc.write(os);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally{
			closeStream(is, os);
		}
	}
	
	/**
	 * 读写word中表格的单元格文本。
	 */
	@Test
	public void tableTest() {
		//ps: 用代码添加表格及行列的话，格式会出问题，有时候word还会坏掉。
		
		String inputPath = "C:/Users/30616/Desktop/temp/test.doc";
		String outputPath = "C:/Users/30616/Desktop/temp/testResult.doc";
		InputStream is = null;
		OutputStream os = null;
		
		try {
			is = new FileInputStream(inputPath);
			HWPFDocument doc = new HWPFDocument(is);
			Range range = doc.getRange();
			
			TableIterator it = new TableIterator(range);
			while (it.hasNext()) {
				Table table = (Table) it.next();
				
				for (int rowIdx = 0; rowIdx < table.numRows(); rowIdx++) {
					TableRow row = table.getRow(rowIdx);

					for (int colIdx = 0; colIdx < row.numCells(); colIdx++) {
						TableCell cell = row.getCell(colIdx);

						Paragraph par = cell.getParagraph(0);

						System.out.println(par.text());
						par.insertBefore("" + (rowIdx * row.numCells() + colIdx));

					}
				}
				
			}
			
			os = new FileOutputStream(outputPath);
			doc.write(os);

		} catch (Exception e) {
			e.printStackTrace();
		} finally{
			closeStream(is, os);
		}
	}
	
	/**
	 * 这是我写的工具类，可以先替换占位符，然后填充表格数据（表格中的开始位置可选）。数据见参数map。
	 * 
	 *  如果word模板中存在表格，在表格的第一个单元格写$table1#1#2#3#4:等参数，
	 *  其中table1为了对应参数，1234（从1开始计数）代表数据填充从表格第1-2行，第3-4列开始，
	 *  也可以不指定行列数，写$table1:或者$table1#d#d#3#4:，其中d代表默认，既从第2行(通常第1行是标题)第1列开始，写数据至表格上限。
	 */
	@Test
	public void poiUtilsTest(){
		
		String resPath = "C:/Users/30616/Desktop/temp/test.doc";
		String outPath = "C:/Users/30616/Desktop/temp/testResult.doc";
		
		
		HashMap<String, Object> params = new HashMap<String, Object>();
		
		HashMap<String, String> mainParams = new HashMap<String, String>();
		mainParams.put("aaa", "啊啊啊");
		
		params.put("main", mainParams);
		
		HashMap<String, String> table1Params = new HashMap<String, String>();
		table1Params.put("tb1_zzz1", "z1");
		table1Params.put("tb1_zzz2", "z2");
		table1Params.put("tb1_zzz3", "z3");
		
		ArrayList<HashMap<String, String>> tableList = new ArrayList<HashMap<String, String>>();
		tableList.add(table1Params);
		tableList.add(table1Params);
		tableList.add(table1Params);
	 
		params.put("table1", tableList);
		
		try {
			POIUtils.docParamConvert(resPath, outPath, params);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	/**
	 * 关闭流
	 * 
	 * @param is
	 * @param os
	 */
	private static void closeStream(InputStream is, OutputStream os) {
		if (is != null) {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		if (os != null) {
			try {
				os.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	
	public static void main(String[] args) {
		// 不知道为什么，用test就会报错，用main就没问题。
		POIXwpfTest test = new POIXwpfTest();
		
//		test.wordReplaceTest();
//		test.tableTest();
		test.poiUtilsTest();
		
	}
	
}


























