package com.kuani.excel;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;


import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@SuppressWarnings({"java:S106","java:S4823","java:S1192"})
public final class DeliveryNote {
	
	public DeliveryNote(String file_path) {
		
		setFile_Path(file_path);
		
	}
	
	public static String file_Path;
	
	
	public static String getFile_Path() {
		return file_Path;
	}

	public static void setFile_Path(String file_Path) {
		DeliveryNote.file_Path = file_Path;
	}

	public static final String[] titles = {
            "品名",   "數量", "", "箱數","箱", "P/O NO", "棧板"," "," "," "
    };

    public static Object[][] getEpsData() {
		return eps_data;
	}

	public static void setEpsData(Object[][] epsData) {
		eps_data = epsData;
	}

	
	public static Object[][] eps_data;
    
	public static String getTimeNumber() {
        String pattern = "yyyyMMdd";
        SimpleDateFormat d = new SimpleDateFormat(pattern);
        return d.format(new Date());
    }
	
	
    //處理外箱麥頭編排方式
    public String procStr(String str) {
    	
        	String mastr ="";
        	char[] chars = str.toCharArray();
        	int ent = 0;
        	
        	for(int i=0;i<chars.length;i++) {
//        		System.out.println("char("+i+")"+chars[i]+" ascii="+ (int)chars[i] );
        	      if((int)chars[i] == 32 && ent == 32) {
        	    	  mastr += chars[i]+"\n";
        	    	  ent = 0;
        	      } else if((int)chars[i] == 32) {
        	    	  ent = 32;
         	      } else 
        			  mastr += chars[i];
        		} 
        	
        	return mastr;
    		}
    	

    public boolean generateXLS(String sheetName, HashMap _epsmap) {
        boolean flag = false;

        Workbook wb = new XSSFWorkbook();
        Map<String, CellStyle> styles = createStyles(wb);
        //處理若資料超過25自動存取在新分頁
        int batchSize = 25;
        int totalDataSize = eps_data.length;

        for (int start = 0; start < totalDataSize; start += batchSize) {
            int end = Math.min(start + batchSize, totalDataSize);
            Object[][] batchData = Arrays.copyOfRange(eps_data, start, end);

            Sheet sheet = wb.createSheet(sheetName + "_" + (start / batchSize + 1));
            

            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setLandscape(false);
            sheet.setFitToPage(true);
            sheet.setHorizontallyCenter(true);
            

            //space row        
            Row spaceRow = sheet.createRow(0);
            spaceRow.setHeightInPoints(5);
            Cell spaceCell = spaceRow.createCell(0);
            spaceCell.setCellValue(" ");
            spaceCell.setCellStyle(styles.get("space")); 
            
            //title row
            Row titleRow = sheet.createRow(1);
            titleRow.setHeightInPoints(40);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue("冠億齒輪股份有限公司");
            titleCell.setCellStyle(styles.get("title"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("$A$2:$J$2"));
            
            titleRow = sheet.createRow(2);
            titleRow.setHeightInPoints(40);
            titleCell = titleRow.createCell(0);
            titleCell.setCellValue("送貨單");
            titleCell.setCellStyle(styles.get("title"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("$A$3:$J$3"));
            
            //General data 
            spaceRow = sheet.createRow(3);
            spaceRow.setHeightInPoints(40);
            spaceCell = spaceRow.createCell(0);
            spaceCell.setCellValue((String) _epsmap.get("customer"));
            spaceCell.setCellStyle(styles.get("highline"));
            
            spaceCell = spaceRow.createCell(2);
            spaceCell.setCellValue("台中市大甲區幼三路3號");
            spaceCell.setCellStyle(styles.get("space"));
            
            spaceCell = spaceRow.createCell(5);
            spaceCell.setCellValue("統編:36108173");
            spaceCell.setCellStyle(styles.get("space"));
            
            spaceCell = spaceRow.createCell(8);
            spaceCell.setCellValue((String) _epsmap.get("shipno"));
            spaceCell.setCellStyle(styles.get("highline"));
            
            spaceRow = sheet.createRow(4);
            spaceRow.setHeightInPoints(40);
            
            spaceCell = spaceRow.createCell(2);
            spaceCell.setCellValue("TEL:04-26812249");
            spaceCell.setCellStyle(styles.get("space"));
            
            spaceCell = spaceRow.createCell(5);
            spaceCell.setCellValue("FAX:04-26817399");
            spaceCell.setCellStyle(styles.get("space"));
            
            spaceCell = spaceRow.createCell(8);
            spaceCell.setCellValue((String)_epsmap.get("shipdate"));
            spaceCell.setCellStyle(styles.get("space"));   
            
            //header row
            Row headerRow = sheet.createRow(5);
            headerRow.setHeightInPoints(35);
            Cell headerCell;
            for (int i = 0; i < titles.length; i++) {
            	if(i != 1) {
            		headerCell = headerRow.createCell(i);
                    headerCell.setCellValue(titles[i]);
                    headerCell.setCellStyle(styles.get("header_center"));
            	} else {
            		headerCell = headerRow.createCell(i);
                    headerCell.setCellValue(titles[i]);
                    headerCell.setCellStyle(styles.get("header_center"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("$B$6:$C$6"));
            	}
            }
            
            headerCell = headerRow.createCell(7);
            headerCell.setCellValue("送貨地點:");
            headerCell.setCellStyle(styles.get("header"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("H6:H8"));
            
            
            headerCell = headerRow.createCell(8);
            headerCell.setCellValue((String)_epsmap.get("shiplocale"));
            headerCell.setCellStyle(styles.get("right_23"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("I6:J8"));
                 
            
          
             int rownum = 6;
             
            for (int i = 0; i < 29; i++) {
                Row row = sheet.createRow(rownum++);
                row.setHeightInPoints(40);
                
                for (int j = 0; j < titles.length; j++) {
                	if(j == 0) {
                		
                		Cell cell = row.createCell(j);               
                        cell.setCellStyle(styles.get("cell_left"));
                    
                	} else {
                       Cell cell = row.createCell(j);               
                       cell.setCellStyle(styles.get("cell"));
                	}
                   
                }
                
                if(i == 2) {
                	
                	//System.out.println("rownum="+rownum);
                	
                	Cell soCell;
                
                	soCell = row.createCell(7);
                	soCell.setCellValue("S/O NO:");
                	soCell.setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H9:H10"));
                    
                    soCell = row.createCell(8);
                    soCell.setCellValue((String)_epsmap.get("sono"));
                    soCell.setCellStyle(styles.get("right_23"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("I9:J10"));
                	
                }
                if(i == 4) {
                	Cell rightCell;
                    
                	rightCell = row.createCell(7);
                	rightCell.setCellValue("船名航次:");
                	rightCell.setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H11:H14"));
                    
                    rightCell = row.createCell(8);
                    rightCell.setCellValue((String)_epsmap.get("voyage"));
                    rightCell.setCellStyle(styles.get("right_23"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("I11:J14"));
                	
                }
                if(i == 8) {
                	Cell rightCell;
                    
                	rightCell = row.createCell(7);
                	rightCell.setCellValue("報關行:");
                	rightCell.setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H15:H16"));
                    
                    rightCell = row.createCell(8);
                    rightCell.setCellValue((String)_epsmap.get("broker"));
                    rightCell.setCellStyle(styles.get("right_23"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("I15:J16"));
                	
                }
                if(i == 10) {
                	Cell rightCell;
                    
                	rightCell = row.createCell(7);
                	rightCell.setCellValue("TEL");
                	rightCell.setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H17:H18"));
                    
                    rightCell = row.createCell(8);
                    rightCell.setCellValue((String)_epsmap.get("tel"));
                    rightCell.setCellStyle(styles.get("right_23"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("I17:J18"));
                	
                }
                if(i == 12) {
                	Cell rightCell;
                    
                	rightCell = row.createCell(7);
                	rightCell.setCellValue("聯絡人:");
                	rightCell.setCellStyle(styles.get("header"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H19:H20"));
                    
                    rightCell = row.createCell(8);
                    rightCell.setCellValue((String)_epsmap.get("contact"));
                    rightCell.setCellStyle(styles.get("right_23"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("I19:J20"));
                	
                }
                if(i == 14) {
                	Cell rightCell;
                    
                	rightCell = row.createCell(7);
                	rightCell.setCellValue("外箱麥頭\n"+procStr((String)_epsmap.get("ma_header")));
                	rightCell.setCellStyle(styles.get("header_left"));
                    sheet.addMergedRegion(CellRangeAddress.valueOf("H21:J32"));
                   
                	
                }
            }          
            headerRow = sheet.createRow(32);
            headerRow.setHeightInPoints(35);
            headerCell = headerRow.createCell(0);
            headerCell.setCellValue("");
            headerCell.setCellStyle(styles.get("right_23"));
            sheet.addMergedRegion(CellRangeAddress.valueOf("A33:G34"));
            
            headerCell = headerRow.createCell(7);
            headerCell.setCellValue("車行/司機");
            headerCell.setCellStyle(styles.get("header"));
            
            headerCell = headerRow.createCell(8);
            headerCell.setCellValue(" ");
            headerCell.setCellStyle(styles.get("right_23"));
            
            headerCell = headerRow.createCell(9);
            headerCell.setCellValue((String)_epsmap.get("tot_weight"));
            headerCell.setCellStyle(styles.get("right_23"));
            
            headerRow = sheet.createRow(33);
            headerCell = headerRow.createCell(7);
            headerCell.setCellValue("車號");
            headerCell.setCellStyle(styles.get("header"));
            
            headerCell = headerRow.createCell(8);
            headerCell.setCellValue(" ");
            headerCell.setCellStyle(styles.get("right_23"));
            
            headerCell = headerRow.createCell(9);
            headerCell.setCellValue((String)_epsmap.get("volume"));
            headerCell.setCellStyle(styles.get("right_23"));
            
            
            headerRow = sheet.createRow(34);
            headerRow.setHeightInPoints(21);
            for(int i=0;i<9; i++) {
            	if(i==0) {
            		headerCell = headerRow.createCell(0);
                    headerCell.setCellValue("審核:");
                    headerCell.setCellStyle(styles.get("footer"));
            	} else if(i == 4) {
            		headerCell = headerRow.createCell(4);
                    headerCell.setCellValue("製表人 : ");
                    headerCell.setCellStyle(styles.get("footer"));       		
            	} else if(i == 7) {
            		headerCell = headerRow.createCell(7);
                    headerCell.setCellValue("出貨人 : ");
                    headerCell.setCellStyle(styles.get("footer"));
            	} else if(i==8) {
            		headerCell = headerRow.createCell(9);
                    headerCell.setCellValue("守衛 :");
                    headerCell.setCellStyle(styles.get("footer"));
            	} else {
            		headerCell = headerRow.createCell(i);
                    headerCell.setCellValue(" ");
                    headerCell.setCellStyle(styles.get("footer"));
            	}
            	  
            
            }
          
          //set sample data
            for (int i = 0; i < batchData.length; i++) {
            	Row row = sheet.getRow(6 + i);
            	 for (int j = 0; j < eps_data[i].length; j++) {
                     if (batchData[i][j] == null) continue;

//                     if (j == 1 || j == 3) {
//                         row.getCell(j).setCellValue(Integer.parseInt((String) batchData[i][j]));
//                     } else {
//                         row.getCell(j).setCellValue((String) batchData[i][j]);
//                     }
                     if (j == 1 || j == 3) {
                    	    // 如果是 "數量" 或 "箱數" 欄位
                    	    row.getCell(j).setCellValue(Double.parseDouble((String) batchData[i][j]));
                    	} else {
                    	    row.getCell(j).setCellValue((String) batchData[i][j]);
                    	}
                 }
            }
            

            Row row = sheet.getRow(31);
            
            for (int j = 0; j < 7; j++) {
            	
            	if(j == 0)
            	   row.getCell(j).setCellValue((String) "合計");
            	 
            	 if(j == 1) 
            		row.getCell(j).setCellFormula("SUM(B7:B31)");
            	
            	 if(j ==2)
            		row.getCell(j).setCellValue((String) "P'S");
            	
            	 if(j ==3)
            		row.getCell(j).setCellFormula("SUM(D7:D31)");
            	 
            	 if(j ==4)
             		row.getCell(j).setCellValue((String) "箱");
            	 
            	 if(j ==5)
              		row.getCell(j).setCellValue((String) "棧板共");
            	 
            	 if(j ==6)
              		row.getCell(j).setCellValue((String) "板");
             	 
            	
            	  row.getCell(j).setCellStyle(styles.get("formula"));
            	
            	
            }
            
            
            //finally set column widths, the width is measured in units of 1/256th of a character width
            //設定欄寬 
            sheet.setColumnWidth(0, 22*420); //30 characters wide
            for (int i = 2; i < 9; i++) {
                sheet.setColumnWidth(i, 10*420);  //6 characters wide
            }
            sheet.setColumnWidth(5, 24*350);
            sheet.setColumnWidth(7, 12*300);
            sheet.setColumnWidth(8, 16*300);
//            sheet.setColumnWidth(9, 22*220);
            sheet.setColumnWidth(9, 22*300);
            //sheet.setColumnWidth(10, 10*256); //10 characters wide
            try {
                String fileName = "DeliveryNote_%1$s.xls";
                fileName = String.format(getFile_Path() + fileName, getTimeNumber());

                if (wb instanceof XSSFWorkbook) fileName += "x";
                FileOutputStream out = new FileOutputStream(fileName);
                wb.write(out);
                out.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return flag;
    }

    public static void main(String[] args) throws Exception {
    	DeliveryNote _delivernot = new DeliveryNote("D:/");
    	
    	//測試資料匯入
    	Object[][] eps_data = {
    			  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "18", "箱", "22042073/4500493491", "1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
                  {"L900620(KI-1480)", "56", "P'S", "7", "箱", "22042073/4500493491", "3"},
                  {"L101142(KI-1838A-6)", "36", "P'S", "17", "箱", "22042073/4500493491","1"},
                  {"L900620(KI-1480)", "14", "P'S", "7", "箱", "22042073/4500493491", "2"},
   
    	};

        DeliveryNote.setEpsData(eps_data);

        HashMap<String, String> _epsmap = new HashMap<String, String>();
        // Your existing map data

    	
    	_epsmap.put("customer", "DATACOLS"); //客戶名
    	_epsmap.put("shipno","NO:111080007"); //送貨單號
    	_epsmap.put("shipdate", "出貨日期:111年8月3日"); //送貨日期
    	_epsmap.put("shiplocale", "貿聯企業股份有限公司"); //送貨地
    	_epsmap.put("sono", "N328"); // S/O No. 
    	_epsmap.put("voyage","AL MURABBA/022W"); //航次
    	_epsmap.put("broker","秉富報關行"); //報關行
    	_epsmap.put("tel","02-2556-0655"); //聯絡電話
    	_epsmap.put("contact","盧小姐.柯小姐"); //聯絡人
    	String str = "　　　╱╲  　　╱　　╲  　╱　　　　╲  <　  AIMSAK　　>  　╲　　　　╱  　　╲　　╱  　　　╲╱ ";
    	   	
    	_epsmap.put("ma_header",str); //麥頭內容
    	_epsmap.put("tot_weight" ,"189'"); //總重量
    	_epsmap.put("volume","2018");  //總材積
        _delivernot.generateXLS("DATACOL", _epsmap);
    }

    /**
     * Create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb){
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style;
        
        Font spaceFont = wb.createFont();
        spaceFont.setFontHeightInPoints((short)15);
        spaceFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(spaceFont);
        styles.put("space", style);
        
        Font highlineFont = wb.createFont();
        highlineFont.setFontHeightInPoints((short)28);
        highlineFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(highlineFont);
        styles.put("highline", style);
        
        Font titleFont = wb.createFont();
        titleFont.setFontHeightInPoints((short)35);
        titleFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(titleFont);
        styles.put("title", style);

        Font monthFont = wb.createFont();
        monthFont.setFontHeightInPoints((short)15);
        //monthFont.setColor(IndexedColors.WHITE.getIndex());
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        monthFont.setBold(true);
        //style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        //style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(monthFont);
        style.setWrapText(true);
        styles.put("header", style);
        
        Font hdFont = wb.createFont();
        hdFont.setFontHeightInPoints((short)15);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        hdFont.setBold(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(hdFont);
        style.setWrapText(true);
        styles.put("header_center", style);
        
        Font headerFont = wb.createFont();
        headerFont.setFontHeightInPoints((short)15);
        style = wb.createCellStyle();
        //style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        headerFont.setBold(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(headerFont);
        style.setWrapText(true);
        styles.put("header_left", style);

 
        Font cellFont = wb.createFont();
        cellFont.setFontHeightInPoints((short)18);
        cellFont.setBold(true);
        style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        hdFont.setBold(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(hdFont);
        style.setWrapText(true);
        styles.put("cell", style);
        
        
        Font cell_leftFont = wb.createFont();
        cell_leftFont.setFontHeightInPoints((short)18);
        cell_leftFont.setBold(true);
        style = wb.createCellStyle();
//        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        hdFont.setBold(true);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(hdFont);
        style.setWrapText(true);
        styles.put("cell_left", style);

        style = wb.createCellStyle();
        //style.setAlignment(HorizontalAlignment.CENTER);        
        Font font_25 = wb.createFont();
        font_25.setFontHeightInPoints((short)20);
        font_25.setBold(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(font_25);
        style.setWrapText(true);
        styles.put("right_23", style);
        
        style = wb.createCellStyle();
        //style.setAlignment(HorizontalAlignment.CENTER);
        Font font_25_Bold = wb.createFont();
        font_25_Bold.setFontHeightInPoints((short)15);
        font_25_Bold.setBold(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setFont(font_25_Bold);
        style.setWrapText(true);
        styles.put("right_25_Bold", style);
        
        style = wb.createCellStyle();
        Font formulaFont = wb.createFont();
        formulaFont.setFontHeightInPoints((short)15);
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        formulaFont.setBold(true);
        style.setFont(formulaFont);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);      
        //style.setDataFormat(wb.createDataFormat().getFormat("0.00"));
        styles.put("formula", style);
        
       
        style = wb.createCellStyle();
        Font footerFont = wb.createFont();
        footerFont.setFontHeightInPoints((short)15);
        style.setFont(footerFont);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());       
        styles.put("footer", style);
        
      

        return styles;
    }
}



