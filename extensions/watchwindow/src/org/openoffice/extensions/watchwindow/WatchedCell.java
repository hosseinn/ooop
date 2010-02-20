package org.openoffice.extensions.watchwindow;

import com.sun.star.container.XNamed;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.document.XEventListener;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.sheet.ComplexReference;
import com.sun.star.sheet.FormulaToken;
import com.sun.star.sheet.SingleReference;
import com.sun.star.sheet.XFormulaParser;
import com.sun.star.sheet.XFormulaTokens;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.CellAddress;
import com.sun.star.table.XCell;
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
    private FormulaToken         m_cellPos                 = null;
    private XCell                m_xCell                   = null ;
    private short                m_num                     = 0;
    private static int           _counter                  = 0;
    private XSpreadsheetDocument m_xDocument               = null;
    private XModifyBroadcaster   m_xCellModifyBroadcaster  = null;
    private XModifyBroadcaster   m_xDocModifyBroadcaster   = null;
    private XEventBroadcaster    m_xEventBroadcaster       = null;
    private String               m_modifySheetName         = "";
   
 
    public WatchedCell(Controller controller, XSpreadsheet xSpreadsheet, String sheetName, FormulaToken cellPos, String selectedCellName, short numer) {
        try {

            m_Controller = controller;
            m_xSpreadsheet = xSpreadsheet;
            m_sheetName = m_modifySheetName = sheetName;
            m_cellPos = cellPos;
            m_cellName = selectedCellName;
            m_num = numer;

            init();
            adjustCounter();
            getController().addToListBox(this.toString(), m_num);

        } catch (IndexOutOfBoundsException ex) {
            ex.printStackTrace();
        }
    }
    
    public void init() throws IndexOutOfBoundsException {

        //m_xCell is the basic of WatchedCell objects. This won't be changed!
        SingleReference ref = (SingleReference) m_cellPos.Data;
        m_xCell = m_xSpreadsheet.getCellByPosition(ref.Column, ref.Row);

        setValue();
        setFormula();

        //add listeners
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

        setCellName();
        setSheetName();
        setValue();
        setFormula();

        if(m_num>=0){
            getController().removeFromListBox(m_num, (short)1);
            getController().addToListBox(this.toString(), m_num);
        }
    }

    public void setCellName(){
        m_cellPos = getController().getRefToken(m_xCell);
        SingleReference ref = (SingleReference)m_cellPos.Data;
        // We don't want to display the sheet name, and the cell position should be absolute, not relative.
        ref.Flags = 0;
        FormulaToken[] tokens = { m_cellPos };
        XFormulaParser xParser = m_Controller.getFormulaParser();
        m_cellName = xParser.printFormula(tokens, new CellAddress());
    }
   
    public String getCellName(){   
        return m_cellName;
    }

    public String getCellNameWithSpaces(){
        return getController().createStringWithSpace(getCellName(), 15);
    }

    public void setSheetName(){
        SingleReference ref = (SingleReference)m_cellPos.Data;
        XSpreadsheet sheet = m_Controller.getSheetByIndex(ref.Sheet);
        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, sheet);
        m_sheetName = xNamed.getName();
    }

    public String getSheetName(){
        return m_sheetName;
    }

    public String getSheetNameWithSpaces(){
        return getController().createStringWithSpace(getSheetName(), 27);
    }

    public void setValue() {
        XText xText = (XText)UnoRuntime.queryInterface(XText.class, m_xCell);
        m_value = xText.getString();
    }

    public String getValue(){
        return m_value;
    }

    public String getValueWithSpaces(){
        return getController().createStringWithSpace(getValue(), 26);
    }

    public void setFormula() {
        XFormulaTokens xTokens = (XFormulaTokens) UnoRuntime.queryInterface(XFormulaTokens.class, m_xCell);
        FormulaToken[] tokens = xTokens.getTokens();
        for(int i=0;i<tokens.length;i++){
            if(tokens[i].Data.toString().startsWith("com.sun.star.sheet.ComplexReference")){
                ComplexReference ref = (ComplexReference) (tokens[i].Data);
                ref.Reference1.Flags = ref.Reference2.Flags = 0;
            }
            if(tokens[i].Data.toString().startsWith("com.sun.star.sheet.SingleReference")){
                SingleReference ref = (SingleReference)(tokens[i].Data);
                ref.Flags = 0;
            }
        }

        m_formula = m_Controller.getFormulaParser().printFormula(tokens, new CellAddress());
        if (!m_formula.isEmpty()) {
            m_formula = "=" + m_formula;
        }    
    }


    public String getFormula(){
        return m_formula;
    }

    public FormulaToken getCellPos(){
        return m_cellPos;
    }

    public void setNum(short n){
        m_num = n;
    }

    public short getNum(){
        return m_num;
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