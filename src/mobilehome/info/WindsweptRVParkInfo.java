package mobilehome.info;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCellFormat;
import jxl.format.Border;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WindsweptRVParkInfo {
	private int LotNumber;
	private float ExpectedMonthlyRent;
	private float StartMeterReading;
	private float EndMeterReading;
	private float ElectricityRate;
	public final static int WINDSWEPT_MAX_LOTS=49; //48 is max but excel has header row
	public WindsweptRVParkInfo(int lotnumber,
								float expectedmonthlyrent,
								float startmeterreading,
								float endmeterreading,
								float electricityrate){
		LotNumber = lotnumber;
		ExpectedMonthlyRent	= expectedmonthlyrent;
		StartMeterReading = startmeterreading;
		EndMeterReading = endmeterreading;
		ElectricityRate = electricityrate;		
	}
	
	public String toString(){
		return "LotNum: " + LotNumber + ": ExpectedMonthlyRent: " + ExpectedMonthlyRent + ", StartMeterReading: " + StartMeterReading
				+ ", EndMeterReading: " + EndMeterReading
				+ ", ElectricityRate: " + ElectricityRate;
	}

	public float getExpectedMonthlyRent() {
		return ExpectedMonthlyRent;
	}
	public void setExpectedMonthlyRent(float expectedMonthlyRent) {
		ExpectedMonthlyRent = expectedMonthlyRent;
	}
	public float getStartMeterReading() {
		return StartMeterReading;
	}
	public void setStartMeterReading(float startMeterReading) {
		StartMeterReading = startMeterReading;
	}
	public float getEndMeterReading() {
		return EndMeterReading;
	}
	public void setEndMeterReading(float endMeterReading) {
		EndMeterReading = endMeterReading;
	}
	public float getElectricityRate() {
		return ElectricityRate;
	}
	public void setElectricityRate(float electricityRate) {
		ElectricityRate = electricityRate;
	}
	public int getLotNumber() {
		return LotNumber;
	}
	public void setLotNumber(int lotNumber) {
		LotNumber = lotNumber;
	}
	public float ElectricityDues(){
		return (getEndMeterReading()-getStartMeterReading())*getElectricityRate();
	}
	
	public float TotalDue(){
		return getExpectedMonthlyRent() + ElectricityDues();
	}

	public static void GenerateRentInvoice(List<WindsweptRVParkInfo> mobilehomeinfo) throws BiffException, IOException, RowsExceededException, WriteException{
		File inputWorkbook = new File("resources/Windswept/invoice.xls");
		File outputWorkbook = new File("resources/Windswept/Windswept_Rent_Invoice.xls");
		//System.out.println(System.getProperty("user.dir"));
		Workbook w1 = Workbook.getWorkbook(inputWorkbook);
	    WritableWorkbook w2 = Workbook.createWorkbook(outputWorkbook, w1);
		
	    WritableFont arial8pt = new WritableFont(WritableFont.ARIAL,8);
	    WritableCellFormat textFormat = new WritableCellFormat (arial8pt);
	    textFormat.setIndentation(1);
	    
		for(WindsweptRVParkInfo mh : mobilehomeinfo){
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
		    Label l = new Label (RentInvoiceTemplate.INVOICE_NO_COL,
			    				RentInvoiceTemplate.INVOICE_NO_ROW,
			    				sdf.format(now)+ "01" + mh.getLotNumber());
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Lot Number
		    
		    //{{SET TITLE
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_TITLE_COL, RentInvoiceTemplate.INVOICE_TITLE_ROW);
		    WritableCellFormat cellFormat = new WritableCellFormat(readCell.getCellFormat());
		    cellFormat.setBorder(Border.TOP,BorderLineStyle.THIN,Colour.GREY_25_PERCENT);
		    cellFormat.setBorder(Border.BOTTOM,BorderLineStyle.THIN,Colour.GREY_25_PERCENT);
			l = new Label (RentInvoiceTemplate.INVOICE_TITLE_COL,
							RentInvoiceTemplate.INVOICE_TITLE_ROW,"WINDSWEPT RV PARK",cellFormat);
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
		    l = new Label (RentInvoiceTemplate.INVOICE_ADDRESS_1ST_COL,
		    					RentInvoiceTemplate.INVOICE_ADDRESS_1ST_ROW,"105 Burch Circle");
		    
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Address Field
		    
		    //{{Set Email Address
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_EMAIL_COL, 
		    						RentInvoiceTemplate.INVOICE_EMAIL_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_EMAIL_COL,
		    					RentInvoiceTemplate.INVOICE_EMAIL_ROW,"Email: windsweptrvpark@yahoo.com");
		    
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Email Address

		    //{{Set Customer ID Number
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_CUST_ID_COL, RentInvoiceTemplate.INVOICE_CUST_ID_ROW);
		    l = new Label (RentInvoiceTemplate.INVOICE_CUST_ID_COL,
			    			RentInvoiceTemplate.INVOICE_CUST_ID_ROW,
			    			"Space " + mh.getLotNumber());
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set Customer ID Number
		    
		    //{{Set To Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_TO_COL, RentInvoiceTemplate.INVOICE_TO_ROW);
			l = new Label (RentInvoiceTemplate.INVOICE_TO_COL,
			    			RentInvoiceTemplate.INVOICE_TO_ROW,
			    			"Space #" + mh.getLotNumber() +" Windswept");
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(l);
		    //}}Set To Field
		    
		    NumberFormat dollarFormat = new NumberFormat(NumberFormat.CURRENCY_DOLLAR + "#,###.00", NumberFormat.COMPLEX_FORMAT);
		    WritableCellFormat wcf = new WritableCellFormat(dollarFormat);
	
		    //{{Set Previous Balance Field
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_COL, RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_ROW);
		    Number n = new Number(RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_COL, 
		    					  RentInvoiceTemplate.INVOICE_PREVIOUS_AMOUNT_DUE_ROW,
		    					  mh.getExpectedMonthlyRent(),wcf);
		    n.setCellFormat(readCell.getCellFormat());
		    sheet.addCell(n);
		    //}}Set Previous Balance Field
	    
		    //{{Set Check-Payable-To
		    readCell = sheet.getCell(RentInvoiceTemplate.INVOICE_ANNOUNCE_COL, 
		    					RentInvoiceTemplate.WSRVP_INVOICE_ANNOUNCE_ROW);
		    String announce=readCell.getContents();
		    l = new Label (RentInvoiceTemplate.INVOICE_ANNOUNCE_COL,
		    			RentInvoiceTemplate.WSRVP_INVOICE_ANNOUNCE_ROW,
		    			announce);
		    l.setCellFormat(readCell.getCellFormat());
		    sheet.setRowView(RentInvoiceTemplate.WSRVP_INVOICE_ANNOUNCE_ROW, 171*20);
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
