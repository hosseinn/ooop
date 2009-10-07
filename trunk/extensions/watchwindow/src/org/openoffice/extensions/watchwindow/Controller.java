package org.openoffice.extensions.watchwindow;

import com.sun.star.awt.Point;
import com.sun.star.container.XNamed;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.CellAddress;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.table.XColumnRowRange;
import com.sun.star.table.XTableColumns;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class Controller {

    private   XComponentContext      m_xContext          = null;
    private   XFrame                 m_xFrame            = null;
    private   XController            m_xController       = null;
    private   Gui                    m_Gui               = null;
    private   DataModel              m_DataModel         = null;
    private   XModel                 m_xModel            = null;
    private   XMultiComponentFactory m_xServiceManager   = null;
    private   XSpreadsheetDocument   m_xDocument         = null;
    private   String                 m_selectedCellName  = null;
    private   String                 m_selectedSheetName = null;
    private   XSpreadsheet           m_xValidSheet       = null;
    private static short             _numer              = 0;
   
    public Controller(XComponentContext xContext, XFrame xFrame){
        m_xContext = xContext;
        m_xFrame = xFrame;
        m_xController = m_xFrame.getController();
        m_xModel = m_xController.getModel();
        m_xServiceManager = m_xContext.getServiceManager();
        m_xDocument = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, m_xModel);      
    }
    
    public XSpreadsheetDocument getDocument(){
        return m_xDocument;
    }
    
    public XModel getModel(){
        return m_xModel;
    }
    
    public Locale getLocation() { 
        Locale locale = null;
        try {
            Object oConfigurationProvider = m_xServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XLocalizable xLocalizable = (XLocalizable) UnoRuntime.queryInterface(XLocalizable.class, oConfigurationProvider);
            locale = xLocalizable.getLocale();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return locale;
    }
    
    public short getNumer(){
        return _numer;
    }
    
    public void decreaseNumer(){
        _numer--;
    }
    
    public void increaseNumer(){
        _numer++;
    }
    
    public void displayWatchWindow() throws Exception{
        if(m_Gui == null){
           m_Gui = new Gui(this, m_xContext, getModel());
        }
        m_Gui.setVisible(true);          
    }
    
    public void addToListBox( String s, short num){
        m_Gui.addToListBox( s, num);
    }
    
    public void removeFromListBox( short num, short n){
        m_Gui.removeFromListBox(num, n);
    }
    
    public void removeSheetItemsFromDataList(String sheetName) { 
        m_DataModel.removeSheetItemsFromDataList(sheetName);
    }
    
    public XSpreadsheet getActiveSheet(){
        XController xController = m_xFrame.getController();
        XSpreadsheetView xView = (XSpreadsheetView)UnoRuntime.queryInterface(XSpreadsheetView.class, xController); 
        return xView.getActiveSheet();  
    } 
    
    public XCell getActiveCell(){
        Object obj = m_xModel.getCurrentSelection();
        return (XCell)UnoRuntime.queryInterface(XCell.class, obj);
    } 
    
    public String getSheetName(XSpreadsheet xSheet){
        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, xSheet);
        return xNamed.getName();
    } 
    
    public String getCellNameWithDollars(XCell xCell) throws Exception{
        String cellName ="";
        Point address = getCellAdress(xCell);
        cellName = "$" + convertColumnIndexToName(address.X);
        cellName += "$" + (address.Y+1);
        return cellName;
    }
    
    public Point getCellAdress(XCell xCell) throws Exception {
        XCellAddressable xCellAddress =  (XCellAddressable)UnoRuntime.queryInterface(XCellAddressable.class, xCell);
        CellAddress aAddress = xCellAddress.getCellAddress();
        Point point =  new Point(aAddress.Column,aAddress.Row);
        return point;
    }
    
    public String convertColumnIndexToName(int columnPos) throws IndexOutOfBoundsException, WrappedTargetException, IndexOutOfBoundsException {   
        XCellRange xCellRange = getActiveSheet().getCellRangeByPosition(columnPos, 0, columnPos, 0);           
        XColumnRowRange xColumnRowRange =(XColumnRowRange)UnoRuntime.queryInterface(XColumnRowRange.class, xCellRange);
        XTableColumns xTableColumns = xColumnRowRange.getColumns();              
        Object oColumnName = xTableColumns.getByIndex(0);
        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, oColumnName);
        return xNamed.getName();    
    } 
    
    public String createStringWithSpace(String s, int max){
        char c;
        int n;
        double dn = s.length()*2;
        for(int i=0; i<s.length();i++){
            c = s.charAt(i);
            if(c=='W')
                dn += 1.3333;
            if(c=='w'||c=='Q'||c=='O'||c=='Ó'||c=='Ö'||c=='Õ'||c=='m'||c=='M')
                dn += 0.6666;
            if(c=='A'||c=='R'||c=='U'||c=='Ú'||c=='Ü'||c=='Á'||c=='D'||c=='G'||c=='H'||c=='Û'||c=='N'||c=='C')
                dn += 0.3333;
            if(c=='z'||c=='J'||c=='L'||c=='k'||c=='s'||c=='c')
                dn -= 0.3333;
            if(c=='I'||c=='Í'||c=='t'||c=='r'||c=='f')
                dn -= 0.6666;
            if(c==' '||c=='j')
                dn -= 1.0;
            if(c=='i'||c=='l'||c=='í'||c=='|'||c=='\'')
                dn -= 1.3333;
        }
        n = (int)Math.round(dn);
        while(n<max){
            s += ' ';
            n++;
        }
        return s;
    }
    
    public boolean isValidSelectedName(String selectAreaName){
        int length = selectAreaName.length();
        if(length>5){
            m_selectedCellName  = "";
            m_selectedSheetName = "";
            short index = 0;
            short m = 0;
            char[] selectedCellCharName = selectAreaName.toCharArray();
            if(index>=length || selectedCellCharName[index++]!='$')
                return false;
            while(index<length && selectedCellCharName[index]!='.' ){
                m_selectedSheetName += selectedCellCharName[index];
                m++;
                index++;
            }  
            if(m==0 || index>=length || selectedCellCharName[index++]!='.')
                return false;
            if(index>=length || selectedCellCharName[index++]!='$')
                return false;
            m = 0;
            while(index<length && selectedCellCharName[index]!='$' && (int)selectedCellCharName[index]>64 && (int)selectedCellCharName[index]<91) {
                m_selectedCellName += selectedCellCharName[index];
                m++;
                index++;
            }
            if(m==0 || index>=length || selectedCellCharName[index++]!='$')
                return false;
            m = 0;
            while(index<length && (int)selectedCellCharName[index]>47 && (int)selectedCellCharName[index]<58 ) {
                m_selectedCellName += selectedCellCharName[index];
                m++;
                index++;
            }
            if(m==0)
                return false;
            Object oSheet = null;
            try{
                XSpreadsheets xSpreadsheets = m_xDocument.getSheets();
                oSheet = xSpreadsheets.getByName(m_selectedSheetName);
                if(oSheet == null )
                    return false; 
            }catch(Exception ex){
                ex.printStackTrace();
                return false;
            }  
            m_xValidSheet = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, oSheet);
             if(m_xValidSheet != null )
                 return true;
        }
        return false;
    }
    
    public void addCell() {
        try {
            String cellName = "";
            XSpreadsheet activeSheet = getActiveSheet();
            XCell activeCell = getActiveCell();
            cellName = "$" + getSheetName(activeSheet) + "." + getCellNameWithDollars(activeCell);
            m_Gui.cellSelection(cellName);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
          
    }
    
    public void removeCell()  {
        short i = m_Gui.getSelectedListBoxPos();
        if(i < 0){
            m_Gui.showMessageBox(1);
        }else{
            //remove i. element into the list
            m_Gui.removeFromListBox(i , (short)1 );     
            m_DataModel.removeFromDataList(i);
            //decrease static member also
            _numer--;
        }
    }

    public void done(String selectAreaName) throws Exception {
            if(isValidSelectedName(selectAreaName)){
                if(m_DataModel == null)
                    m_DataModel = new DataModel(this);
                m_DataModel.addToDataList( m_xValidSheet, m_selectedSheetName, m_selectedCellName);
                m_Gui.setVisible(true);
            }else{
                m_Gui.setVisible(true);
                m_Gui.showMessageBox(0);
            }         
    }
    
}