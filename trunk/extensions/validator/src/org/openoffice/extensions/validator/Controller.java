package org.openoffice.extensions.validator;

import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNamed;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.sheet.ComplexReference;
import com.sun.star.sheet.FormulaToken;
import com.sun.star.sheet.ReferenceFlags;
import com.sun.star.sheet.SingleReference;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XFormulaParser;
import com.sun.star.sheet.XFormulaTokens;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.table.CellAddress;
import com.sun.star.table.CellContentType;
import com.sun.star.table.CellRangeAddress;
import com.sun.star.table.XCell;
import com.sun.star.table.XCellRange;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import java.util.ArrayList;
import java.util.HashSet;

import com.sun.star.lang.Locale;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.text.XText;

public class Controller {

    private XComponentContext       m_xContext          = null;
    private XFrame                  m_xFrame            = null;
    private XSpreadsheetDocument    m_xDocument         = null;

    private Gui                     m_Gui               = null;
    private Model                   m_Model             = null;
    private XSpreadsheet            m_xSpreadsheet      = null;
    private XFormulaParser          m_xFormulaParser    = null;
    private short                   m_num                = 0;


    Controller(XComponentContext xContext, XFrame xFrame) throws Exception {
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_xDocument = (XSpreadsheetDocument) UnoRuntime.queryInterface(XSpreadsheetDocument.class, m_xFrame.getController().getModel());
        if(m_Model == null)
           m_Model = new Model( this );
        if(m_Gui == null)
           m_Gui = new Gui( this, m_xContext, m_xFrame, m_xDocument );
    }

    protected Model getModel(){
        return m_Model;
    }

    protected Gui getGui(){
        return m_Gui;
    }

    public short getNum(){
        return m_num;
    }

    public void setNum( short n){
        m_num = n;
    }

    public void setVisibleSelectWindow( boolean b ){
        m_Gui.setVisibleSelectWindow( b );
    }

    public void setActiveSheet() {
        XSpreadsheetView xView = (XSpreadsheetView)UnoRuntime.queryInterface( XSpreadsheetView.class, m_xFrame.getController() );
        m_xSpreadsheet = xView.getActiveSheet();
    }
    
    public XSpreadsheet getActiveSheet() {
        setActiveSheet();
        return m_xSpreadsheet;
    }

    public XSpreadsheet getClosingActiveSheet() {
        return m_xSpreadsheet;
    }

    public String getActiveSheetName(){
        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, getActiveSheet());
        return xNamed.getName();
    }

    public XSpreadsheet getSheetByIndex( int index ){
        XSpreadsheet xSheet = null;
        try {
            XIndexAccess indexAccess = (XIndexAccess) UnoRuntime.queryInterface( XIndexAccess.class, m_xDocument.getSheets() );
            xSheet = (XSpreadsheet) UnoRuntime.queryInterface( XSpreadsheet.class, indexAccess.getByIndex( index ) );
        }catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in controller.getSheetByIndex(()" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in controller.getSheetByIndex(()" + ex.getLocalizedMessage());
        }
        return xSheet;
    }

    public String getSheetNameByIndex( int index ){
        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, getSheetByIndex(index));
        return xNamed.getName();
    }

    public void addErrorCell( XCell xCell, short errorType, ArrayList<XCell> list ){
        if( list.size() > 0 ) {
            Cell cell = new Cell( m_Gui, m_num, xCell, errorType, list );
            getModel().addItemToList( cell );
            getGui().addItemToListBox( cell.toString(), m_num );
            m_num++;
        }
    }

    public void setErrorList() {
        try {
            XCellRangesQuery xCellQuery = (XCellRangesQuery)UnoRuntime.queryInterface( XCellRangesQuery.class, getActiveSheet() );
            XSheetCellRanges xFormulaCells = xCellQuery.queryContentCells( (short)com.sun.star.sheet.CellFlags.FORMULA );
            XEnumerationAccess xFormulas = xFormulaCells.getCells();
            XEnumeration xFormulaEnum = xFormulas.createEnumeration();
            Object formulaCell = null;

            //  observ all formula cells
            while ( xFormulaEnum.hasMoreElements() ) {
                formulaCell = xFormulaEnum.nextElement();
                XCell xCell = (XCell)UnoRuntime.queryInterface( XCell.class, formulaCell );
                //XText xText = (XText)UnoRuntime.queryInterface(XText.class, xCell);
                //System.out.println("XCellName: " + getCellName( xCell ) + " XCellFormula: " + xCell.getFormula() +" XCellType: " + xCell.getType().getValue() + " XCellError: " + xCell.getError());

                // if it is not usual (predefinied) error
                int errorNum = xCell.getError();

                if( errorNum == 0 ) {

                    ArrayList<XCell> lPrecedentsCells = getPrecedentCellsOfFormulaCell(xCell);
                    // formulas without precedents (as =234)
                    if( lPrecedentsCells == null ){
                        //System.out.println("null formula");
                    } else {
                        if( getGui().m_errorType1 ) {
                            ArrayList<XCell> list = getError1List( lPrecedentsCells ) ;
                            if( list.size() > 0 )
                                addErrorCell( xCell, (short)1, getError1List( list ) );
                        }
                        if( getGui().m_errorType2 ) {
                            ArrayList<XCell> list = getError2List( lPrecedentsCells );
                            if( list.size() > 0 )
                                addErrorCell( xCell, (short)2, list );
                        }
                    }
                } else {
                    if( errorNum == 503 ||  errorNum == 532 ){
                        ArrayList<XCell> list = getError503List( xCell );
                        addErrorCell( xCell, (short)503, list );
                    }
                }
            }
        } catch (Exception ex) {
            System.err.println("Exception in Gui.setErrorCells(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public void setError3List() {
        try {
            HashSet<String> allPrecedentNames = getPrecedentOfAllFormulas();
            XCellRangesQuery xCellQuery = (XCellRangesQuery)UnoRuntime.queryInterface(XCellRangesQuery.class, getActiveSheet());
            XSheetCellRanges xValueCells = xCellQuery.queryContentCells( (short)com.sun.star.sheet.CellFlags.VALUE );
            XEnumerationAccess xValues = xValueCells.getCells();
            XEnumeration xValueEnum = xValues.createEnumeration();
            Object valueCell = null;
            boolean isContain; // true (valid) if at least a formula contains the cell
            // observ all valueCells
            while ( xValueEnum.hasMoreElements() ) {
                isContain = false;
                valueCell = xValueEnum.nextElement();
                XCell xValueCell = (XCell) UnoRuntime.queryInterface( XCell.class, valueCell );
                String currValueCellName = getCellName(xValueCell);
                //XText xText = (XText)UnoRuntime.queryInterface(XText.class, xValueCell);
                //System.out.println("XCellName: " + currValueCellName + " XCellFormula: " + xValueCell.getFormula() +" XCellType: " + xValueCell.getType().getValue() + " XCellError: " + xValueCell.getError() +" text: " + xText.getString());
                if( !allPrecedentNames.isEmpty() )
                    for (String cellName : allPrecedentNames)
                        if (cellName.equals(currValueCellName))
                            isContain = true;
                // if the valueCell is not a precedent of any formula, it is not valid
                if ( !isContain ) {
                    ArrayList<XCell> list = getError3List( xValueCell );
                    addErrorCell( xValueCell, (short)3, list );
                }
            }
        } catch (NoSuchElementException ex) {
            System.err.println("NoSuchElementException in Controller.setError3List(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in Controller.setError3List(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    // Return a list wich contains emty value cells of formula cell.
    // If list is empty, there is not any empty value cell in precedents of formula of xCell.
    // It returns null if the xCell formula is not valid.
    public ArrayList<XCell> getError1List( ArrayList<XCell> lPrecedentsCells ){
        if( lPrecedentsCells == null)
            return null;
        ArrayList<XCell> list = new ArrayList<XCell>();
        for(XCell currXCell : lPrecedentsCells)
            if(currXCell.getType().getValue()== CellContentType.EMPTY_value)
                list.add(currXCell);
        return list;
    }
 
    public ArrayList<XCell> getError2List( ArrayList<XCell> lPrecedentsCells ){
        if( lPrecedentsCells == null)
            return null;
        ArrayList<XCell> list = new ArrayList<XCell>();
        for(XCell currXCell : lPrecedentsCells)
            if(currXCell.getType().getValue()==CellContentType.TEXT_value)
                list.add(currXCell);
        return list;
    }

    public ArrayList<XCell> getError3List(XCell xCell){
        boolean isNumber = false;
        boolean isContain = false; // true (valid) if at least a formula contains the cell
        ArrayList<XCell> list = new ArrayList();

        HashSet<String> allPrecedentNames = getPrecedentOfAllFormulas();
        if( xCell.getType().getValue() == CellContentType.VALUE_value ) 
            isNumber = true;
        String cellName = getCellName( xCell );
        if( !allPrecedentNames.isEmpty() )
            for ( String currCellName : allPrecedentNames )
                if ( cellName.equals( currCellName ) )
                    isContain = true;
        if ( isNumber && !isContain )
            list.add(xCell);
        return list;
    }

    public ArrayList<XCell> getError503List( XCell xCell ){
        if( xCell.getError() == 503 || xCell.getError() == 532 ){
            ArrayList<XCell> list = getPrecedentCellsOfFormulaCell(xCell);
            if(list == null){
                list = new ArrayList<XCell>();
                list.add(xCell);
            }
            return list;
        }else
            return null;
    }

    public Locale getLocation() {
        Locale locale = null;
        try {
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object oConfigurationProvider = xMCF.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XLocalizable xLocalizable = (XLocalizable) UnoRuntime.queryInterface(XLocalizable.class, oConfigurationProvider);
            locale = xLocalizable.getLocale();
        } catch (Exception ex) {
            System.err.println("Exception in Gui.getLocation(). Message:\n" + ex.getLocalizedMessage());
        }
        return locale;
    }
/*
    public String conversSpecificFNameToOriginalFName(String formula){
        String newformula = formula.toUpperCase();
        String lang = getLocation().Language;

        if(!lang.equals("en")){

            String[] aFunctionNames = null;

            if(lang.equals("hu"))
               aFunctionNames = huFunctionNames;

            for(int i = 0; i < aFunctionNames.length; i++)
                if(!aFunctionNames[i].equals(""))
                    if(newformula.contains(aFunctionNames[i]))
                        newformula = newformula.replaceAll(aFunctionNames[i], enFunctionNames[i]);

        }
        return newformula;
    }

    public String conversOriginalMNameToSpecificMName(String methodName){
        String newMethodName = methodName.toUpperCase();
        String lang = getLocation().Language;
        System.out.println(lang);

        if(!lang.equals("en")){

            String[] aFunctionNames = null;
            if(lang.equals("hu"))
               aFunctionNames = huFunctionNames;

            for(int i = 0; i < enFunctionNames.length; i++)
                if(newMethodName.contains(enFunctionNames[i]))
                    if(!aFunctionNames[i].equals(""))
                        newMethodName = newMethodName.replaceAll(enFunctionNames[i], aFunctionNames[i]);

        }
        return newMethodName;
    }
*/
    public HashSet<String> getPrecedentOfAllFormulas(){

        HashSet<String> precedentNames = new HashSet<String>();

        XCellRangesQuery xCellQuery = (XCellRangesQuery)UnoRuntime.queryInterface( XCellRangesQuery.class, getActiveSheet() );
        XSheetCellRanges xFormulaCells = xCellQuery.queryContentCells( (short)com.sun.star.sheet.CellFlags.FORMULA );
        XEnumerationAccess xFormulas = xFormulaCells.getCells();
        XEnumeration xFormulaEnum = xFormulas.createEnumeration();
        Object formulaCell = null;

        while ( xFormulaEnum.hasMoreElements() ) {
            try {
                formulaCell = xFormulaEnum.nextElement();
                XCell xCell = (XCell)UnoRuntime.queryInterface( XCell.class, formulaCell );
                XText xText = (XText)UnoRuntime.queryInterface( XText.class, xCell );
                if(!(xCell.getError() == 524  && xText.getString().equals("#REF!"))){
                    ArrayList<XCell> lPrecedentsCells = getPrecedentCellsOfFormulaCell(xCell);
                    if( lPrecedentsCells != null )
                        for(XCell currXCell : lPrecedentsCells)
                            precedentNames.add( getCellName( currXCell ) );
                }
            } catch (Exception ex) {
                System.err.println("Exception in Gui.setErrorCells(). Message:\n" + ex.getLocalizedMessage());
            }
        }
        return precedentNames;
    }

    // return with the errorType
    // 0 - no error in list
    // 1  - there is empty cell in list
    // 2  - there is string value cell in list
    public short parseXCells(ArrayList<XCell> list){
        for( XCell currXCell : list ){
            if(currXCell.getType().getValue()== CellContentType.EMPTY_value)
               return 1;
            if(currXCell.getType().getValue()== CellContentType.TEXT_value)
               return 2;
        }
        return 0;
    }

    // return the precedent cells of xCell
    public ArrayList<XCell> getPrecedentCellsOfFormulaCell(XCell xCell){
        XFormulaTokens xTokens = (XFormulaTokens) UnoRuntime.queryInterface(XFormulaTokens.class, xCell);
        FormulaToken[] tokens = null;
        tokens = xTokens.getTokens();
        return parseFormula(tokens);
    }

    // return the precedent cells of formula
    public ArrayList<XCell> getPrecedentCellsOfFormula(String formula){
        FormulaToken[] tokens = null;
        tokens = getFormulaParser().parseFormula(formula, new CellAddress());
        return parseFormula(tokens);
    }

    public String getCellName(XCell xCell) {
        FormulaToken formulaToken = getRefToken(xCell);
        SingleReference ref = (SingleReference)formulaToken.Data;
        // We don't want to display the sheet name, and the cell position should be absolute, not relative.
        ref.Flags = 0;
        FormulaToken[] tokens = { formulaToken };
        XFormulaParser xParser = getFormulaParser();
        XCellAddressable xCellAddr =  (XCellAddressable)UnoRuntime.queryInterface(XCellAddressable.class, xCell);
        CellAddress addr = xCellAddr.getCellAddress();
        return xParser.printFormula(tokens, addr);
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

    public XFormulaParser getFormulaParser() {
        if (m_xFormulaParser == null) {
            try {
                // We need to get a service factory from the desktop instance
                // in order to instantiate the formula parser.
                Object oDesktop = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", m_xContext);
                XDesktop xDesktop = (XDesktop)UnoRuntime.queryInterface(XDesktop.class, oDesktop);
                XMultiServiceFactory xSrvMgr = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, xDesktop.getCurrentComponent());
                Object oParser = xSrvMgr.createInstance("com.sun.star.sheet.FormulaParser");
                m_xFormulaParser = (XFormulaParser)UnoRuntime.queryInterface(XFormulaParser.class, oParser);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return m_xFormulaParser;
    }

    public String getFormula(XCell xCell) {
        String formula = "";
        XFormulaTokens xTokens = (XFormulaTokens) UnoRuntime.queryInterface(XFormulaTokens.class, xCell);
        FormulaToken[] tokens = xTokens.getTokens();
        XCellAddressable xCellAddr =  (XCellAddressable)UnoRuntime.queryInterface(XCellAddressable.class, xCell);
        CellAddress addr = xCellAddr.getCellAddress();
        formula = getFormulaParser().printFormula( tokens, addr );
        if ( !formula.isEmpty() )
            formula = "=" + formula;
        return formula;
    }

    // return the precedents of formulaToken
    // it give null when the formula (formulaToken) is not valid (not formula)
    // getPrecedentCellsOfFormula methods use this to parse
    public ArrayList<XCell> parseFormula(FormulaToken[] tokens){
        ArrayList<CellRangeAddress> lAddresses = new ArrayList<CellRangeAddress>();
        ArrayList<XCell> lPrecedentsCells = null;
        for(int i=0; i<tokens.length; i++) {
            if(tokens[i].Data.toString().startsWith("com.sun.star.sheet.ComplexReference")){
                ComplexReference ref = (ComplexReference) tokens[i].Data;
                ref.Reference1.Flags = ref.Reference2.Flags = 0;

                CellRangeAddress address = new CellRangeAddress();
                address.Sheet = (short)ref.Reference1.Sheet;
                address.StartColumn = ref.Reference1.Column;
                address.StartRow    = ref.Reference1.Row;
                address.EndColumn   = ref.Reference2.Column;
                address.EndRow      = ref.Reference2.Row;
                lAddresses.add( address );
            }
            if(tokens[i].Data.toString().startsWith("com.sun.star.sheet.SingleReference")){
                SingleReference ref = (SingleReference)tokens[i].Data;
                ref.Flags = 0;

                CellRangeAddress address = new CellRangeAddress();
                address.Sheet = (short)ref.Sheet;
                address.StartColumn = address.EndColumn = ref.Column;
                address.StartRow    = address.EndRow = ref.Row;
                lAddresses.add( address);
            }
        }
        if( lAddresses.size() > 0 ){
            lPrecedentsCells        = new ArrayList<XCell>();
            XCellRange xCellRange   = null;
            XCell currXCell         = null;
            try {
                for(CellRangeAddress address : lAddresses) {
                    xCellRange = getSheetByIndex(address.Sheet).getCellRangeByPosition(address.StartColumn, address.StartRow, address.EndColumn, address.EndRow);
                    for (int ii = 0; ii < address.EndColumn - address.StartColumn + 1; ii++) {
                        for (int j = 0; j < address.EndRow - address.StartRow + 1; j++) {
                            currXCell = xCellRange.getCellByPosition(ii, j);
                            lPrecedentsCells.add(currXCell);
                        }
                    }
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println("IndexOutOfBoundsException in controller.parseFormula()" + ex.getLocalizedMessage());
            }
        }
        return lPrecedentsCells;
    }
    
}
