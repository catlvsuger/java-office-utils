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
 * ����Apache-POI
 *  ʵ�ֶ�word�ĵ��Ĵ���
 * @author shizhongyu
 */
@SuppressWarnings("unchecked")
public class POIUtils {
	
	/**
	 * ��ȡģ��word�������еĲ������滻��ֵ�����������ļ����
	 *  ֧�ֵ�word��ʽ��03�� doc
	 *  
	 * @param resPath 	ģ��·��
	 * @param outPath 	���·��
	 * @param params 	��Ҫ������ʽ�� HashMap< String : HashMap<String, String> >
	 * 					��������ʽ�� HashMap< String : List< HashMap<String, String> > >
	 * 					{"main" : { key_1:value, key_2:value }, 
	 * 					 "table1" : { { key1_1:value, key1_2:value }, { key2_1:value,key2_2:value  } }
	 * @throws Exception 
	 * @tips �轫word�е�ͼƬ��ʽ���ó�"Ƕ���Ͱ�ʽ"
	 * @tips key��Ӧword�е�${key}
	 */
	public static void docParamConvert(String resPath, String outPath,
			HashMap<String, Object> params) throws Exception {

		InputStream is = null;
		OutputStream os = null;
		try {
			is = new FileInputStream(resPath);
			HWPFDocument doc = new HWPFDocument(is);
			Range range = doc.getRange();
			
			/** �����Ҫ�������������ģ� */
			POIUtils.replaceMain(range, params);
			
			/** ���������� */
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
	 * �����Ҫ�������������ģ�
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
	 * ����������
	 * @param range
	 * @param params
	 */
	private static void replaceTables(Range range, Map params){
		//�����ĵ��е����б��
		TableIterator it = new TableIterator(range);
		while (it.hasNext()) {
			Table table = (Table) it.next();
		    
		    //�ҵ���һ��Ҫ�滻�����ı�����滻����
		    String[] tableOpt = obtainTableName(table);
		    if(tableOpt == null){
		    	//��ǰ�������  ������Ҫ�滻�ı��	��continue
		    	continue;
		    }
		    
		    //ȷ�������Ҫ�滻�ķ�Χ��x�е�x�У�x�е�x��(��0��ʼ)
		    int[] replaceRange = obtainReplaceRange(table, tableOpt);
		    
		    //�ӱ������Ҫ�滻�Ĳ�����
		    String[] paramsOrder = obtainParamsOrder(table, replaceRange);
		    if(paramsOrder == null){
		    	continue;
		    }
		    
		    //�����滻
		    List<Map<String, String>> tableParam = (List<Map<String, String>>) params.get(tableOpt[0]); //��map����table��Ӧ���滻����
		    replaceTable(tableParam, paramsOrder, replaceRange, table, range);
		    
		}
		
	}
	/**
	 * �����滻
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
	 * ȡ���е�0��cell������ $tableX#1#2#3#4:zzzz ��ʽ��
	 *  ���滻Ϊzzzz��������tableX��1 2 3 4�����򲻲�����������null��
	 * @param table
	 * @return tableOpt ����[0] �����滻�ӱ�ĵ�[1]�е���[2]�У���[3]�е���[4]��
	 */
	private static String[] obtainTableName(Table table){
	    Paragraph tableSign =  table.getRow(0).getCell(0).getParagraph(0);
	    String text = tableSign.text();
	    String[] texts = text.split(":");	//texts[0]�����	texts[1]ԭ��д�������text
	    if(texts.length < 2){
	    	return null;
	    }
	    String tabletext = texts[0].trim();
	    if(tabletext.indexOf("table") == -1){
	    	return null;
	    }
	    tableSign.replaceText(text, texts[1]);//���˾�ɾ��
	    
	    String[] tableOpt = tabletext.split("#");
	    for (int i = 0; i < tableOpt.length; i++) {
	    	tableOpt[i] = tableOpt[i].trim();
		}
	    tableOpt[0] = tableOpt[0].substring(1, tableOpt[0].length()); //�����
	    
	    return tableOpt;
	}
	/**
	 * �ӱ������Ҫ�滻�Ĳ�����
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
            //replaceText��ʱ�򣬼ǵ�����word�еĽ�������
            par.replaceText( cellText, cellText.substring(cellText.length()-1, cellText.length()) );
            paramsOrder[i] = cellText.substring(2, cellText.length() - 2);

        }
		return paramsOrder;
	}
	/**
	 * ȷ�������Ҫ�滻�ķ�Χ��x[0]�е�x[1]�У�x[2]�е�x��[3](��0��ʼ)
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
	 * �ر���
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
		mainParams.put("aaa", "������");
		
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














