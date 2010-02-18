package org.openoffice.extensions.watchwindow;

import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.sheet.XFormulaParser;
import com.sun.star.sheet.FormulaToken;
import com.sun.star.sheet.ReferenceFlags;
import com.sun.star.sheet.SingleReference;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.CellAddress;
import com.sun.star.table.XCell;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class Controller {

    private   XComponentContext      m_xContext          = null;
    private   XFrame                 m_xFrame            = null;
    private   XController            m_xController       = null;
    protected   Gui                    m_Gui               = null;
    private   DataModel              m_DataModel         = null;
    private   XModel                 m_xModel            = null;
    private   XMultiComponentFactory m_xServiceManager   = null;
    private   XSpreadsheetDocument   m_xDocument         = null;
    private   XFormulaParser         m_xFormulaParser    = null;
    private   FormulaToken           m_selectedCellPos   = null;
    private   XSpreadsheet           m_xValidSheet       = null;
    private static short             _numer              = 0;

    //members of constants group AddressConvention, have not yet used
    //private static final short      OOO                  = 0;
    //private static final short      XL_A1                = 1;
    //private static final short      XL_R1C1              = 2;
    //private static final short      XL_OOX               = 3;
    //private static final short      LOTUS_A1             = 4;


   
    public Controller(XComponentContext xContext, XFrame xFrame) throws UnknownPropertyException, WrappedTargetException, IllegalArgumentException{
        m_xContext = xContext;
        m_xFrame = xFrame;
        m_xController = m_xFrame.getController();
        m_xModel = m_xController.getModel();
        m_xServiceManager = m_xContext.getServiceManager();
        m_xDocument = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, m_xModel);
    }

    public XFormulaParser getFormulaParser() {
        if (m_xFormulaParser == null) {
            try {
                // We need to get a service factory from the desktop instance
                // in order to instantiate the formula parser.
                Object oDesktop = m_xServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", m_xContext);
                XDesktop xDesktop = (XDesktop)UnoRuntime.queryInterface(XDesktop.class, oDesktop);
                XMultiServiceFactory xSrvMgr = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, xDesktop.getCurrentComponent());
                Object oParser = xSrvMgr.createInstance("com.sun.star.sheet.FormulaParser");
                m_xFormulaParser = (XFormulaParser)UnoRuntime.queryInterface(XFormulaParser.class, oParser);
                
                //adjust the style of referencia with AddressConvention constants group
                //XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xFormulaParser);
                //xProp.setPropertyValue("FormulaConvention", OOO);

            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return m_xFormulaParser;
    }

    protected FormulaToken getRefToken(XCell xCell) {
        XCellAddressable xCellAddr =  (XCellAddressable)UnoRuntime.queryInterface(XCellAddressable.class, xCell);
        CellAddress addr = xCellAddr.getCellAddress();
        SingleReference ref = new SingleReference();
        ref.Column = addr.Column;
        ref.Row    = addr.Row;
        ref.Sheet  = addr.Sheet;
        ref.Flags = ReferenceFlags.SHEET_3D;
        FormulaToken token = new FormulaToken();
        token.OpCode = 0;
        token.Data = ref;
        return token;
    }

    protected XSpreadsheet getSheetByIndex(int index) {
        if (m_xDocument == null)
            return null;

        XIndexAccess xIA = (XIndexAccess)UnoRuntime.queryInterface(XIndexAccess.class, m_xDocument.getSheets());
        if (index >= xIA.getCount()) {
            return null;
        }
        try {
            XSpreadsheet xSheet = (XSpreadsheet)UnoRuntime.queryInterface(XSpreadsheet.class, xIA.getByIndex(index));
            return xSheet;
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return null;
    }


    public XComponentContext getContext(){
        return m_xContext;
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
           m_Gui = new Gui(this, m_xContext, m_xModel);
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
 
    public String createStringWithSpace(String s, int max){
        char c;
        int n;
        double dn = s.length()*2;
        for(int i=0; i<s.length();i++){
            c = s.charAt(i);
            if(c=='W')
                dn += 1.3333;
            if(c=='w'||c=='Q'||c=='O'||c=='�'||c=='�'||c=='�'||c=='m'||c=='M')
                dn += 0.6666;
            if(c=='A'||c=='R'||c=='U'||c=='�'||c=='�'||c=='�'||c=='D'||c=='G'||c=='H'||c=='�'||c=='N'||c=='C')
                dn += 0.3333;
            if(c=='z'||c=='J'||c=='L'||c=='k'||c=='s'||c=='c')
                dn -= 0.3333;
            if(c=='I'||c=='�'||c=='t'||c=='r'||c=='f')
                dn -= 0.6666;
            if(c==' '||c=='j')
                dn -= 1.0;
            if(c=='i'||c=='l'||c=='�'||c=='|'||c=='\'')
                dn -= 1.3333;
        }
        n = (int)Math.round(dn);
        while(n<max){
            s += ' ';
            n++;
        }
        return s;
    }

    //check the input of user and adjust m_xValidSheet, m_selectedCellPos members
    public boolean isValidSelectedName(String selectAreaName){
        XFormulaParser xParser = getFormulaParser();
        CellAddress origin = new CellAddress();
        FormulaToken[] tokens = xParser.parseFormula(selectAreaName, origin);
        if (tokens.length == 0)
            return false;

        FormulaToken token = tokens[0];
        if (token.OpCode != 0)
            // This is undocumented, but a reference token has an opcode of 0.
            // This may change in the future, to probably a named constant
            // value.
            return false;

        if (SingleReference.class != token.Data.getClass())
            // this must be a single cell reference.
            return false;

        SingleReference ref = (SingleReference)token.Data;
        m_xValidSheet = getSheetByIndex(ref.Sheet);
        if (m_xValidSheet == null)
            return false;

        m_selectedCellPos = token;

        return true;
    }

    public void addCell() {
        //default feature: user can choose the active cell in active sheet
        FormulaToken token = getRefToken(getActiveCell());
        FormulaToken[] tokens = { token };
        String cellName = getFormulaParser().printFormula(tokens, new CellAddress());
        m_Gui.cellSelection(cellName);
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
            if (isValidSelectedName(selectAreaName)) {
            if(m_DataModel == null)
                m_DataModel = new DataModel(this);
            m_DataModel.addToDataList( m_xValidSheet, m_selectedCellPos);
            m_Gui.setVisible(true);
        } else {
            m_Gui.setVisible(true);
            m_Gui.showMessageBox(0);
        }
    }
    
}