package com.poi.xwpf;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

/**
 * 调用Apache-POI
 *  实现对word文档的处理
 * @author shizhongyu
 */
@SuppressWarnings("unchecked")
public class POIUtils {
	
	/**
	 * 读取模板word，将其中的参数项替换成值，并生成新文件输出
	 *  支持的word格式：03版 doc
	 *  
	 * @param resPath 	模板路径
	 * @param outPath 	输出路径
	 * @param params 	主要参数格式： HashMap< String : HashMap<String, String> >
	 * 					表格参数格式： HashMap< String : List< HashMap<String, String> > >
	 * 					{"main" : { key_1:value, key_2:value }, 
	 * 					 "table1" : { { key1_1:value, key1_2:value }, { key2_1:value,key2_2:value  } }
	 * @throws Exception 
	 * @tips 需将word中的图片版式设置成"嵌入型版式"
	 * @tips key对应word中的${key}
	 */
	public static void docParamConvert(String resPath, String outPath,
			HashMap<String, Object> params) throws Exception {

		InputStream is = null;
		OutputStream os = null;
		try {
			is = new FileInputStream(resPath);
			HWPFDocument doc = new HWPFDocument(is);
			Range range = doc.getRange();
			
			/** 填充主要参数（表格以外的） */
			POIUtils.replaceMain(range, params);
			
			/** 填充各表格参数 */
			POIUtils.replaceTables(range, params);

			os = new FileOutputStream(outPath);
			doc.write(os);

		} catch (Exception e){
			e.printStackTrace();
		} finally {
			closeStream(is, os);
		}
		
	}
	
	/**
	 * 填充主要参数（表格以外的）
	 * @param range
	 * @param params
	 */
	private static void replaceMain(Range range, Map params){
		HashMap<String, String> mainParam = (HashMap<String, String>) params.get("main");
		
		for (Entry<String, String> entry : mainParam.entrySet()) {
			range.replaceText("${"+entry.getKey().toLowerCase()+"}", entry.getValue());
		}
	}
	
	/**
	 * 填充各表格参数
	 * @param range
	 * @param params
	 */
	private static void replaceTables(Range range, Map params){
		//迭代文档中的所有表格
		TableIterator it = new TableIterator(range);
		while (it.hasNext()) {
			Table table = (Table) it.next();
		    
		    //找到下一个要替换参数的表格及其替换规则
		    String[] tableOpt = obtainTableName(table);
		    if(tableOpt == null){
		    	//当前迭代表格  不是需要替换的表格	则continue
		    	continue;
		    }
		    
		    //确定表格需要替换的范围：x行到x行，x列到x列(从0开始)
		    int[] replaceRange = obtainReplaceRange(table, tableOpt);
		    
		    //从表格中拿要替换的参数名
		    String[] paramsOrder = obtainParamsOrder(table, replaceRange);
		    if(paramsOrder == null){
		    	continue;
		    }
		    
		    //参数替换
		    List<Map<String, String>> tableParam = (List<Map<String, String>>) params.get(tableOpt[0]); //从map中拿table对应的替换参数
		    replaceTable(tableParam, paramsOrder, replaceRange, table, range);
		    
		}
		
	}
	/**
	 * 参数替换
	 * @param tableParam
	 * @param paramsOrder
	 * @param replaceRange
	 * @param table
	 * @param range
	 */
	private static void replaceTable(List<Map<String, String>> tableParam, 
			String[] paramsOrder, int[] replaceRange, Table table, Range range) {
	    
	    for ( int rowIdx = 0; rowIdx < replaceRange[1] - replaceRange[0] + 1; rowIdx++ ){
	        TableRow row = table.getRow( rowIdx + replaceRange[0] );
	        
	        Map<String, String> rowParam = null;
	        try{
	        	rowParam = tableParam.get(rowIdx);
	        }catch(Exception e) {
	        	return;
	        }
	        
	        for ( int colIdx = 0; colIdx < replaceRange[3] - replaceRange[2] + 1; colIdx++ ){
	            TableCell cell = row.getCell( colIdx + replaceRange[2] );
	            
	            Paragraph par = cell.getParagraph( 0 );
	            String param = rowParam.get( paramsOrder[colIdx] );
//	            String param = rowParam.get( paramsOrder[colIdx].toUpperCase() );
	            if(param != null){
            		par.insertBefore(param);
	            }
	        }
	    }
	}
	/**
	 * 取表中第0个cell，若是 $tableX#1#2#3#4:zzzz 格式，
	 *  则替换为zzzz，并返回tableX及1 2 3 4；否则不操作，并返回null。
	 * @param table
	 * @return tableOpt 表名[0] 参数替换从表的第[1]行到第[2]行，第[3]列到第[4]列
	 */
	private static String[] obtainTableName(Table table){
	    Paragraph tableSign =  table.getRow(0).getCell(0).getParagraph(0);
	    String text = tableSign.text();
	    String[] texts = text.split(":");	//texts[0]表格名	texts[1]原来写在上面的text
	    if(texts.length < 2){
	    	return null;
	    }
	    String tabletext = texts[0].trim();
	    if(tabletext.indexOf("table") == -1){
	    	return null;
	    }
	    tableSign.replaceText(text, texts[1]);//拿了就删掉
	    
	    String[] tableOpt = tabletext.split("#");
	    for (int i = 0; i < tableOpt.length; i++) {
	    	tableOpt[i] = tableOpt[i].trim();
		}
	    tableOpt[0] = tableOpt[0].substring(1, tableOpt[0].length()); //表格名
	    
	    return tableOpt;
	}
	/**
	 * 从表格中拿要替换的参数名
	 * @param table
	 * @return
	 */
	private static String[] obtainParamsOrder(Table table, int[] replaceRange){
		String[] paramsOrder = new String[replaceRange[3] - replaceRange[2] + 1];
		
		TableRow paramsRow =  table.getRow(replaceRange[0]);
        for ( int i = 0; i < replaceRange[3] - replaceRange[2] + 1; i++ ){
            TableCell cell = paramsRow.getCell( i + replaceRange[2] );
            
            Paragraph par = cell.getParagraph( 0 );
            
            String cellText = par.text();
            if(cellText.length() == 1){
            	return null;
            }
            //replaceText的时候，记得留下word中的结束符。
            par.replaceText( cellText, cellText.substring(cellText.length()-1, cellText.length()) );
            paramsOrder[i] = cellText.substring(2, cellText.length() - 2);

        }
		return paramsOrder;
	}
	/**
	 * 确定表格需要替换的范围：x[0]行到x[1]行，x[2]列到x列[3](从0开始)
	 * @param table
	 * @param tableOpt
	 * @return
	 */
	private static int[] obtainReplaceRange(Table table, String[] tableOpt) {
		//$table1#1#2 -> $table1#1#2#d#d
		String[] rangeStr = {"d", "d", "d", "d"};	//defualt
		for (int i = 0; i < tableOpt.length - 1; i++) {
			rangeStr[i] = tableOpt[i+1];
		}
		
		int[] range = new int[4];
		for (int i = 0; i < range.length; i++) {
			if ("d".equals(rangeStr[i])) {
				switch (i) {
				case 0:
					range[i] = 1;
					break;
				case 1:
					range[i] = table.numRows() - 1;
					break;
				case 2:
					range[i] = 0;
					break;
				case 3:
					range[i] = table.getRow(0).numCells() - 1;
					break;
				}
			} else {
				range[i] = Integer.parseInt(rangeStr[i]) - 1;
			}
		}
		
		return range;
	}
	
	/**
	 * 关闭流
	 * @param is
	 * @param os
	 */
	private static void closeStream(InputStream is, OutputStream os){
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
	
	public static void main(String[] args) throws Exception {
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
		
		POIUtils.docParamConvert(resPath, outPath, params);
	}

}














