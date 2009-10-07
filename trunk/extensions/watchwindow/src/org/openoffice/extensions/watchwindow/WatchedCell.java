package org.openoffice.extensions.watchwindow;

import com.sun.star.container.XNamed;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.document.XEventListener;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.sheet.XCellRangeAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.text.XText;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XModifyListener;


public class WatchedCell implements XModifyListener, XEventListener{

    private Controller           m_Controller              = null;
    private String               m_sheetName               = "";
    private String               m_cellName                = "";
    private String               m_value                   = "";
    private String               m_formula                 = "";
    private XSpreadsheet         m_xSpreadsheet            = null;
    private XCell                m_xCell                   = null ;
    private short                m_num                     = 0;
    private static int           _counter                  = 0;
    private XSpreadsheetDocument m_xDocument               = null;
    private XModifyBroadcaster   m_xCellModifyBroadcaster  = null;
    private XModifyBroadcaster   m_xDocModifyBroadcaster   = null;
    private XEventBroadcaster    m_xEventBroadcaster       = null;
    private String               m_modifySheetName         = "";
   
    public WatchedCell(Controller controller, XSpreadsheet xSpreadsheet, String sheetName, String cellName, short numer) {
        try {
            m_Controller = controller;
            m_xSpreadsheet = xSpreadsheet;
            m_sheetName = m_modifySheetName = sheetName;
            m_cellName = cellName;
            m_num = numer;
            init();
            adjustCounter();
            getController().addToListBox(this.toString(), m_num);
        } catch (IndexOutOfBoundsException ex) {
            ex.printStackTrace();
        }
    }
    
    public void init() throws IndexOutOfBoundsException {
        XCellRange xCellRange = m_xSpreadsheet.getCellRangeByName(getCellName());
        XCellRangeAddressable xRangeAddr = (XCellRangeAddressable)UnoRuntime.queryInterface(XCellRangeAddressable.class, xCellRange);
        CellRangeAddress cellAddress = xRangeAddr.getRangeAddress();
        m_xCell = m_xSpreadsheet.getCellByPosition(cellAddress.StartColumn, cellAddress.StartRow);
        setValue();
        setFormula(); 
        m_xDocument = getController().getDocument(); 
        m_xDocModifyBroadcaster = (XModifyBroadcaster) UnoRuntime.queryInterface(XModifyBroadcaster.class, m_xDocument);  
        m_xDocModifyBroadcaster.addModifyListener(this);
        m_xEventBroadcaster = (XEventBroadcaster) UnoRuntime.queryInterface(XEventBroadcaster.class, m_xDocument);  
        m_xEventBroadcaster.addEventListener(this);
        m_xCellModifyBroadcaster = (XModifyBroadcaster)UnoRuntime.queryInterface(XModifyBroadcaster.class, m_xCell );
        m_xCellModifyBroadcaster.addModifyListener(this);
    }

    public Controller getController(){
        return m_Controller;
    }
    
    public int countSheets(){
        return m_xDocument.getSheets().getElementNames().length;
    }
    
    public void adjustCounter(){
        _counter = m_xDocument.getSheets().getElementNames().length;
    }
    
    public void refresh(){
        setSheetName();
        setValue();
        setFormula();
        if(m_num>=0){
            getController().removeFromListBox(m_num, (short)1);
            getController().addToListBox(this.toString(), m_num);
        }   
    }
    
    public short getNum(){
        return m_num;
    }
    
    public void setNum(short n){
        m_num = n;
    }
    
    public String getSheetName(){
        return m_sheetName;
    }
    
    public void setSheetName(){
       XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, m_xSpreadsheet);
       m_sheetName = xNamed.getName();
    }
    
    public void setSheetName(String sheetName){
        m_sheetName = sheetName;
    }
    
    public String getSheetNameWithSpaces(){
        return getController().createStringWithSpace(getSheetName(), 27);
    }
    public String getCellName(){
        return m_cellName;
    }
    
    public String getCellNameWithSpaces(){
        return getController().createStringWithSpace(getCellName(), 15);
    }
    
    public void setCellName(String cellName){
        m_cellName = cellName;
    }
    
    public void setValue() { 
        XText xText = (XText)UnoRuntime.queryInterface(XText.class, m_xCell);
        m_value = xText.getString();
    }
    
    public void setValue(String value) {
        m_value = value;
    }
    
    public String getValue(){
        return m_value;
    }
    
    public String getValueWithSpaces(){
        return getController().createStringWithSpace(getValue(), 26);
    }
    public void setFormula(){
        m_formula = m_xCell.getFormula();
        if(!m_formula.equals("")){
            if(m_formula.charAt(0) != '='){
                m_formula ="";
            }
        }
    }
    
    public void setFormula(String formula){
        m_formula = formula;
    }
    public String getFormula(){
        return m_formula;
    }
    
    @Override
    public String toString(){
        return getSheetNameWithSpaces()+getCellNameWithSpaces()+getValueWithSpaces()+getFormula();
    }
    
    //com.sun.star.util.XModifyListener:(2)
    public void modified(EventObject event){
        if(_counter-1 == countSheets()){
            if(m_modifySheetName.equals(getSheetName())){
                getController().removeSheetItemsFromDataList(getSheetName());
                adjustCounter();
           }
        }else{
            adjustCounter();
        }
        refresh();
    }
    
    public void disposing(EventObject event) { 
        m_xCellModifyBroadcaster.removeModifyListener(this);
    }
    
    //com.sun.star.document.XEventListener:
    public void notifyEvent(com.sun.star.document.EventObject event) {
         if(event.EventName.equals("OnModeChanged")){
            XModel xModel = (XModel) UnoRuntime.queryInterface(XModel.class, event.Source);
            XSpreadsheetView xView = (XSpreadsheetView) UnoRuntime.queryInterface(XSpreadsheetView.class, xModel.getCurrentController());
            XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, xView.getActiveSheet());
            m_modifySheetName = xNamed.getName();
         }
    }
    
} 