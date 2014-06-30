package mobilehome.info;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class MobileHomeInfo {
	
	public final static int CT = 1;
	public final static int MESA = 2;
	public final static int WINDSWEPT = 3;
	private float ExpectedMonthlyRent;
	private float PreviousBalance;
	private float LateFee;
	private float Credit;
	private int LotNumber;
	public final static int CT_MAX_LOTS=27; //25 is max but excel has header row
	public final static int MESA_MAX_LOTS=31; //25 is max but excel has header row
	
	public int getLotNumber(){return LotNumber;}
	public float getCredit(){return Credit;}
	public float getLateFee(){return LateFee;}
	public float getPreviousBalance(){return PreviousBalance;}
	public float getExpectedMonthlyRent(){return ExpectedMonthlyRent;}

	public final static int LISTALLHOMES=100;
	
	public static String LotBalanceURL(int n){
		int column = 17; //Column 'Q'
		int minrow = 0;
		int maxrow = 0;
		if(n == 100){
			minrow = 2;
			maxrow = 26;
		}else{
			minrow=1+n;
			maxrow=minrow;
		}
		return "?min-row="+minrow+"&min-col="+column+"&max-row="+maxrow+"&max-col="+column;
	}
	
	public MobileHomeInfo(float expectedmonthlyrent, 
						float previousbalance,
						float latefee,
						float credit,
						int lotnumber){
		ExpectedMonthlyRent = expectedmonthlyrent;
		PreviousBalance = previousbalance;
		LateFee = latefee;
		Credit = credit;
		LotNumber = lotnumber;
	}
	
	public String toString(){
		return "LotNum: " + LotNumber + ": ExpectedMonthlyRent: " + ExpectedMonthlyRent + ", PreviousBalance: " + PreviousBalance
				+ ", LateFee: " + LateFee
				+ ", Credit: " + Credit;
	}
	
	public float TotalDue(){
		return ExpectedMonthlyRent + PreviousBalance + LateFee + Credit;
	}
	
	public static void GenerateRentInvoice(List<MobileHomeInfo> mobilehomeinfo, int mobilepark) throws BiffException, IOException, RowsExceededException, WriteException{
		File inputWorkbook = new File("resources/invoice.xls");
		File outputWorkbook = null;
		if(mobilepark == MobileHomeInfo.CT)
			outputWorkbook = new File("resources/CT_Rent_Invoice.xls");
		else
			outputWorkbook = new File("resources/Mesa_Rent_Invoice.xls");
		//System.out.println(System.getProperty("user.dir"));
		Workbook w1 = Workbook.getWorkbook(inputWorkbook);
	    WritableWorkbook w2 = Workbook.createWorkbook(outputWorkbook, w1);
		
	    WritableFont arial8pt = new WritableFont(WritableFont.ARIAL,8);
	    WritableCellFormat textFormat = new WritableCellFormat (arial8pt);
	    textFormat.setIndentation(1);
	    //System.out.println("Park: " + mobilepark);
	    
		for(MobileHomeInfo mh : mobilehomeinfo){
			if(mh.getExpectedMonthlyRent() == 0) {
				continue;
			}
			String sheetname = "Lot " + mh.getLotNumber();
			w2.copySheet(0, sheetname, mh.getLotNumber());
		    WritableSheet sheet = w2.getSheet(sheetname);
		       
		    //{{Set Lot Number
		    Cell readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_NO_COL, RentInvoiceTemplate.INVOICE_NO_ROW);
		    SimpleDateFormat sdf = new SimpleDateFormat("YYYYMMdd");
		    Date now = new Date();
		    //System.out.println(sdf.format(now));
		    Label l = null;
		    if(mobilepark == MobileHomeInfo.CT){
			    l = new Label (RentInvoiceTemplate.INVOICE_NO_COL,
			    				RentInvoiceTemplate.INVOICE_NO_ROW,
			    				sdf.format(now)+ "06" + mh.getLotNumber());
		    }else{
			    l = new Label (RentInvoiceTemplate.INVOICE_NO_COL,
    							RentInvoiceTemplate.INVOICE_NO_ROW,
    							sdf.format(now)+ "07" + mh.getLotNumber());
		    }
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Lot Number
		    
		    //{{SET TITLE
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_TITLE_COL, 
					RentInvoiceTemplate.INVOICE_TITLE_ROW);
			if(mobilepark == MobileHomeInfo.CT){
				l = new Label (RentInvoiceTemplate.INVOICE_TITLE_COL,
								RentInvoiceTemplate.INVOICE_TITLE_ROW,"CROSS TIMBERS MOBILE HOME PARK");
			}else{
				l = new Label (RentInvoiceTemplate.INVOICE_TITLE_COL,
						RentInvoiceTemplate.INVOICE_TITLE_ROW,"MESA MOBILE HOME PARK");
			}
			
			l.setCellFormat(readCell.getCellFormat());
			sheet.addCell(l);
		    //}}SET TITLE
		    
		    //{{Set Date Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_DATE_COL, RentInvoiceTemplate.INVOICE_DATE_ROW);
		    sdf = new SimpleDateFormat("MMMM d, yyyy");
		    l = new Label (RentInvoiceTemplate.INVOICE_DATE_COL,RentInvoiceTemplate.INVOICE_DATE_ROW,sdf.format(now));
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Date Field
		    
		    //{{Set Address Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_ADDRESS_1ST_COL, 
		    						RentInvoiceTemplate.INVOICE_ADDRESS_1ST_ROW);
		    if(mobilepark == MobileHomeInfo.CT){
		    	l = new Label (RentInvoiceTemplate.INVOICE_ADDRESS_1ST_COL,
		    					RentInvoiceTemplate.INVOICE_ADDRESS_1ST_ROW,"4507 W Oak Street");
		    }else{
		    	l = new Label (RentInvoiceTemplate.INVOICE_ADDRESS_1ST_COL,
    					RentInvoiceTemplate.INVOICE_ADDRESS_1ST_ROW,"1118 N Fort Street");
		    }
		    
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Address Field
		    
		    //{{Set Email Address
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_EMAIL_COL, 
		    						RentInvoiceTemplate.INVOICE_EMAIL_ROW);
		    if(mobilepark == MobileHomeInfo.CT){
		    	l = new Label (RentInvoiceTemplate.INVOICE_EMAIL_COL,
		    					RentInvoiceTemplate.INVOICE_EMAIL_ROW,"Email: crosstimbersmhp@yahoo.com");
		    }else{
		    	l = new Label (RentInvoiceTemplate.INVOICE_EMAIL_COL,
    					RentInvoiceTemplate.INVOICE_EMAIL_ROW,"Email: mesamhp@yahoo.com");
		    }
		    
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Email Address

		    //{{Set Customer ID Number
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_CUST_ID_COL, RentInvoiceTemplate.INVOICE_CUST_ID_ROW);
		    if(mobilepark == MobileHomeInfo.CT){
			    l = new Label (RentInvoiceTemplate.INVOICE_CUST_ID_COL,
			    			RentInvoiceTemplate.INVOICE_CUST_ID_ROW,
			    			"CT Lot " + mh.getLotNumber());
		    }else{
			    l = new Label (RentInvoiceTemplate.INVOICE_CUST_ID_COL,
		    			RentInvoiceTemplate.INVOICE_CUST_ID_ROW,
		    			"Mesa Lot " + mh.getLotNumber());
		    }
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Customer ID Number
		    
		    //{{Set To Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_TO_COL, RentInvoiceTemplate.INVOICE_TO_ROW);
		    if(mobilepark == MobileHomeInfo.CT){
			    l = new Label (RentInvoiceTemplate.INVOICE_TO_COL,
			    			RentInvoiceTemplate.INVOICE_TO_ROW,
			    			"Lot #" + mh.getLotNumber() +" Cross Timbers");
		    }else{
			    l = new Label (RentInvoiceTemplate.INVOICE_TO_COL,
			    			RentInvoiceTemplate.INVOICE_TO_ROW,
			    			"Lot #" + mh.getLotNumber() +" Mesa");
		    }
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set To Field
		    
		    NumberFormat dollarFormat = new NumberFormat(NumberFormat.CURRENCY_DOLLAR + "#,###.00", NumberFormat.COMPLEX_FORMAT);
		    WritableCellFormat wcf = new WritableCellFormat(dollarFormat);
	
		    //{{Set Previous Balance Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_COL, RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_ROW);
		    Number n = new Number(RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_COL, 
		    					  RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_ROW,
		    					  mh.getPreviousBalance(),wcf);
		    n.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(n);
		    //}}Set Previous Balance Field
	
		    //{{Set Late Fee Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_PREVIOUS_LATE_FEE_COL, RentInvoiceTemplate.INVOICE_PREVIOUS_LATE_FEE_ROW);
		    n = new Number(RentInvoiceTemplate.INVOICE_PREVIOUS_LATE_FEE_COL, 
		    			   RentInvoiceTemplate.INVOICE_PREVIOUS_LATE_FEE_ROW,
		    			   mh.getLateFee(),wcf);
		    n.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(n);
		    //}}Set Late Fee Field
	
		    //{{Set Credit Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_CREDIT_COL, RentInvoiceTemplate.INVOICE_CREDIT_ROW);
		    n = new Number(RentInvoiceTemplate.INVOICE_CREDIT_COL, 
		    			   RentInvoiceTemplate.INVOICE_CREDIT_ROW,
		    			   mh.getCredit(),wcf);
		    n.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(n);
		    //}}Set Credit Field
	
		    //{{Set Rent Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_RENT_COL, RentInvoiceTemplate.INVOICE_RENT_ROW);
		    n = new Number(RentInvoiceTemplate.INVOICE_RENT_COL, 
		    			   RentInvoiceTemplate.INVOICE_RENT_ROW,
		    			   mh.getExpectedMonthlyRent(),wcf);
		    n.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(n);
		    //}}Set Rent Field
		    
		    Calendar cal = GregorianCalendar.getInstance();
		    SimpleDateFormat df = new SimpleDateFormat("MMMM");
		    SimpleDateFormat dfYear = new SimpleDateFormat("YYYY");
		    cal.setTime(new Date());
		    cal.add(Calendar.MONTH, 1);
		    String nextMonthAsString = df.format(cal.getTime());
		    String yearAsString = dfYear.format(cal.getTime());
		    l = new Label(RentInvoiceTemplate.INVOICE_RENT_MONTH_COL, 
		    			RentInvoiceTemplate.INVOICE_RENT_ROW,
		    			nextMonthAsString + " " + yearAsString + " Rent");
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_RENT_MONTH_COL,RentInvoiceTemplate.INVOICE_RENT_ROW);
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    
		    //{{Set DUE FIELDS
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_DUE_BEFORE_MONTH_5TH_COL, 
		    		RentInvoiceTemplate.INVOICE_DUE_BEFORE_MONTH_5TH_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_DUE_BEFORE_MONTH_5TH_COL,
		    			RentInvoiceTemplate.INVOICE_DUE_BEFORE_MONTH_5TH_ROW,
		    			"DUE BEFORE " + nextMonthAsString.toUpperCase() + " 5th");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);

		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_5TH_COL, 
		    		RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_5TH_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_5TH_COL,
		    			RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_5TH_ROW,
		    			"DUE AFTER " + nextMonthAsString.toUpperCase() + " 5th");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);

		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_10TH_COL, 
		    		RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_10TH_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_10TH_COL,
		    			RentInvoiceTemplate.INVOICE_DUE_AFTER_MONTH_10TH_ROW,
		    			"DUE AFTER " + nextMonthAsString.toUpperCase() + " 10th");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}
		    
		    //{{Set Late Fee Warning Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_5TH_LATE_FEE_COL, RentInvoiceTemplate.INVOICE_5TH_LATE_FEE_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_5TH_LATE_FEE_COL,
		    			RentInvoiceTemplate.INVOICE_5TH_LATE_FEE_ROW,
		    			"Late Fee After " + nextMonthAsString + " 5th 2:00PM");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_10TH_LATE_FEE_COL, RentInvoiceTemplate.INVOICE_10TH_LATE_FEE_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_10TH_LATE_FEE_COL,
		    			RentInvoiceTemplate.INVOICE_10TH_LATE_FEE_ROW,
		    			"Late Fee After " + nextMonthAsString + " 10th 2:00PM");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Late Fee Warning Field
		    
		    //{{Set Check-Payable-To
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_ANNOUNCE_COL, 
		    					RentInvoiceTemplate.INVOICE_ANNOUNCE_ROW);
		    String announce=readCell.getContents();
		    if(mobilepark == MobileHomeInfo.MESA){
		    	announce=announce.replaceAll("Cross Timbers", "Mesa");
		    	announce=announce.replaceAll("CROSS TIMBERS", "MESA");
		    }
		    l = new Label (RentInvoiceTemplate.INVOICE_ANNOUNCE_COL,
		    			RentInvoiceTemplate.INVOICE_ANNOUNCE_ROW,
		    			announce);
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.setRowView(RentInvoiceTemplate.INVOICE_ANNOUNCE_ROW, 171*20);
			sheet.addCell(l);
		    //}}Set Check-Payable-To
		}//for
	    
		Sheet[] sheets = w2.getSheets();  
		Integer index=null;
		for(int i=0;i<sheets.length && index == null; i++) {
			if(sheets[i].getName().equals("Lot XX")) {
				index = i;
			}
		}
	    w2.removeSheet(index);
	    
	    w2.write();
	    w2.close();		
	}
}
