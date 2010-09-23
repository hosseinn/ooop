package org.openoffice.extensions.validator;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.sheet.SingleReference;
import com.sun.star.table.TableBorder;
import com.sun.star.table.XCell;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XModifyListener;
import java.util.ArrayList;
import java.util.Vector;


public class Cell implements  XModifyListener, Runnable {

    private Gui                     m_Gui                       = null;
    private short                   m_number                    = -1;
    private XCell                   m_xCell                     = null;
    private String                  m_sFormula                  = "";
    private short                   m_ErrorType                 = 0;
    private String                  m_sErrorMessage             = "";
    private ArrayList<XCell>        m_lNotValidPrecCells        = null;
    private Vector<XPropertySet>    m_vProps                    = null;
    private Vector<Integer>         m_vColors                   = null;
    private Vector<TableBorder>     m_vBorders                  = null;
    private XModifyBroadcaster      m_xCellModifyBroadcaster    = null;
    private Thread                  m_runner                    = null;
    private boolean                 m_isRunning                 = false;
    private static short            _selectedItemPos            = -1;

    private final short             ERRORTYPE1                  = 1; // empty cell references in formulas
    private final short             ERRORTYPE2                  = 2; // string value cell references in formuals
    private final short             ERRORTYPE3                  = 3; // value cells which is not used in any formula
    private final short             ERRORTYPE503                = 503; // division by zero
/*
    private final short             ERRORTYPE5                  = 5;
    private final short             ERRORTYPE6                  = 6;
    private final short             ERRORTYPE7                  = 7;
    private final short             ERRORTYPE8                  = 8;
    private final short             ERRORTYPE9                  = 9;
    private final short             ERRORTYPE10                 = 10;
*/


    public Cell ( Gui gui, short number, XCell xCell, short type, ArrayList<XCell> lPrecedentsCells ) {
        this.m_Gui                      = gui;
        this.m_number                   = number;
        this.m_xCell                    = xCell;
        this.m_ErrorType                = type;
        this.m_lNotValidPrecCells       = lPrecedentsCells;
        setErrorMessage();
        if( this.m_ErrorType != ERRORTYPE3 )
            this.m_sFormula             = getGui().getController().getFormula( m_xCell );
        //set precedents of formula cell properties
        setVectors();
        this.m_xCellModifyBroadcaster   = (XModifyBroadcaster)UnoRuntime.queryInterface(XModifyBroadcaster.class, this.m_xCell );
        this.m_xCellModifyBroadcaster.addModifyListener(this);
    }

    public Gui getGui(){
        return m_Gui;
    }

    public XCell getXCell(){
        return this.m_xCell;
    }

    public void setErrorMessage(){
        m_sErrorMessage = getGui().getDialogPropertyValue("Strings", "Strings.errorMessage" + ( m_ErrorType ) + ".Label");
    }

    public String getErrorMessage(){
        return m_sErrorMessage;
    }

    public void clearErrorType(){
        m_ErrorType = 0;
    }

    public void setFormula( String formula){
        m_sFormula = formula;
    }

    public short getNumber(){
        return m_number;
    }

    public ArrayList<XCell> getNotValidPrecCells() {
        return m_lNotValidPrecCells;
    }

    public void setPrecedentsCells(){
        ArrayList<XCell> lAllPrecedentsCells = null;
        clearPrecedentsAndVectors();
        if( m_ErrorType == ERRORTYPE1 || m_ErrorType == ERRORTYPE2 )
            lAllPrecedentsCells = getGui().getController().getPrecedentCellsOfFormulaCell( m_xCell );
        if( m_ErrorType == ERRORTYPE1 )
            m_lNotValidPrecCells = getGui().getController().getError1List( lAllPrecedentsCells );
        if( m_ErrorType == ERRORTYPE2 )
            m_lNotValidPrecCells = getGui().getController().getError2List( lAllPrecedentsCells );
        if( m_ErrorType == ERRORTYPE3 )
            m_lNotValidPrecCells = getGui().getController().getError3List( m_xCell );
        if( m_ErrorType == ERRORTYPE503 )
            m_lNotValidPrecCells = getGui().getController().getError503List(m_xCell);
/*
        if( m_ErrorType == ERRORTYPE5 )
            m_lNotValidPrecCells = getGui().getController().getError5List(m_xCell);
        if( m_ErrorType == ERRORTYPE6 )
            m_lNotValidPrecCells = getGui().getController().getError6List(m_xCell);
        if( m_ErrorType == ERRORTYPE7 )
            m_lNotValidPrecCells = getGui().getController().getError7List(m_xCell);
        if( m_ErrorType == ERRORTYPE8 )
            m_lNotValidPrecCells = getGui().getController().getError8List(m_xCell);
        if( m_ErrorType == ERRORTYPE9 )
            m_lNotValidPrecCells = getGui().getController().getError9List(m_xCell);
        if( m_ErrorType == ERRORTYPE10 )
            m_lNotValidPrecCells = getGui().getController().getError10List(m_xCell);
*/
        setVectors();
    }

    public void clearPrecedentsCellsAndType(){
        clearPrecedentsAndVectors();
        m_ErrorType = 0;
        removeListener();
    }

    public void clearPrecedentsAndVectors(){
        if( m_lNotValidPrecCells != null ){
            m_lNotValidPrecCells.clear();
            m_lNotValidPrecCells = null;
        }
        if( m_vProps != null ){
            m_vProps.clear();
            m_vProps = null;
        }
        if( m_vColors != null){
            m_vColors.clear();
            m_vColors = null;
        }
        if( m_vBorders != null){
            m_vBorders.clear();
            m_vBorders = null;
        }
    }

    public void setVectors(){
        if( m_lNotValidPrecCells != null ){
            m_vProps = new Vector<XPropertySet>();
            m_vColors = new Vector<Integer>();
            m_vBorders = new Vector<TableBorder>();
            if( !m_lNotValidPrecCells.isEmpty() ){
                try {
                    for(XCell xCell : m_lNotValidPrecCells) {
                        XPropertySet xPropSet = (com.sun.star.beans.XPropertySet) UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCell );
                        m_vProps.add(xPropSet);
                        Integer colorValue = (Integer) xPropSet.getPropertyValue( "CellBackColor" );
                        m_vColors.add(colorValue);
                        TableBorder cellBorder= (TableBorder) xPropSet.getPropertyValue( "TableBorder" );
                        m_vBorders.add(cellBorder);

                    }
                } catch (UnknownPropertyException ex) {
                    System.err.println("UnknownPropertyException in Cell.setVectors(). Message:\n" + ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println("WrappedTargetException in Cell.setVectors(). Message:\n" + ex.getLocalizedMessage());
                }
            }
        }
    }
    
    public void setPrecedentsNewColor(){
        try {
            int color = 0xFF0000;
            if( m_ErrorType == ERRORTYPE2 )
                color = 0xFFA500;
            if( m_ErrorType == ERRORTYPE3 )
                color = 0x800000;
            if( m_isRunning && m_lNotValidPrecCells != null && !m_lNotValidPrecCells.isEmpty()){
                TableBorder border = getGui().getTableBorder( new Integer(0x0000FF), (short)75 );
                if( m_vProps != null && !m_vProps.isEmpty() && m_vColors != null && !m_vColors.isEmpty() && m_vBorders != null && !m_vBorders.isEmpty()){
                    for(int i = 0; i < m_vProps.size(); i++){
                        m_vProps.get(i).setPropertyValue( "CellBackColor", new Integer( color ) );
                        m_vProps.get(i).setPropertyValue( "TableBorder", border );
                    }
                }
            } 
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setBackPrecedentsColors(){
        if( m_lNotValidPrecCells != null && !m_lNotValidPrecCells.isEmpty()){
            try {
                if( m_vProps != null && m_vColors != null && m_vBorders != null ) {
                    for(int i = 0; i < m_vProps.size(); i++) {
                        m_vProps.get(i).setPropertyValue("CellBackColor", new Integer(m_vColors.get(i)));
                        m_vProps.get(i).setPropertyValue("TableBorder", m_vBorders.get(i));
                    }
                }
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void removeListener(){
        if( m_xCellModifyBroadcaster != null )
            m_xCellModifyBroadcaster.removeModifyListener(this);
    }

    public short getErrorType(){
        return m_ErrorType;
    }

    public void decreaseNumber(){
        this.m_number--;
    }

    @Override
    public String toString(){
        
        SingleReference formulaCellRef = (SingleReference)getGui().getController().getRefToken(m_xCell).Data;

        String itemValue = getGui().createStringWithSpace( getGui().getController().getCellName( m_xCell ).replace("$", ""), 24 );
        itemValue += getErrorMessage(); // get errorXXX original message, it independent even if the formula had been corrected

        if( m_ErrorType == 0 ) {
            itemValue += " - " + getGui().getDialogPropertyValue("Strings", "Strings.correctedLabel.Label");
        } else {
            if( m_ErrorType != 3 ){
                String sCells = "";
                String currCellName = "";
                if( m_lNotValidPrecCells != null ) {
                    for (XCell currXCell : m_lNotValidPrecCells){
                        SingleReference currCellRef = (SingleReference)getGui().getController().getRefToken(currXCell).Data;
                        currCellName = (currCellRef.Sheet == formulaCellRef.Sheet) ? "" : getGui().getController().getSheetNameByIndex(currCellRef.Sheet) + ".";
                        currCellName += getGui().getController().getCellName(currXCell).replace("$", "");
                        sCells += " " + currCellName;
                    }
                }
                itemValue += " - " + getGui().getDialogPropertyValue("Strings", "Strings.cellName.Label" );
                itemValue += sCells;
            }
        }
        return itemValue;
    }

    public void refreshItemInListBox( short num ){
        getGui().removeItemFromListBox( num );
        getGui().addItemToListBox( toString(), num );
    }
    public void refreshItemInListBoxAndSelect( short num ){
        refreshItemInListBox( num );
        getGui().selectItemPosInListBox( num );
    }

    // com.sun.star.util.XModifyListener (2)
    @Override
    public void modified(EventObject event) {
        _selectedItemPos = getGui().getSelectedItemPosFromListBox();
        
        //  if the user is using the correct option of the Validator
        //  it won't work for the current item in the list box
        if( !getGui().isCorrectProcessRunning() || _selectedItemPos != getNumber() ) {
            stop();
            getGui().clearArrows();
            String sNewFormula = getGui().getController().getFormula( m_xCell );
            //  if the formula of the current xCell is changed to "" (empty string)
            // (not in the correct option of Validator)
            if( !sNewFormula.equals( m_sFormula ) && sNewFormula.equals( "" ) ) {
                setBackPrecedentsColors();
                if( getGui().m_errorType3 && m_ErrorType == ERRORTYPE3 )
                    clearPrecedentsCellsAndType();
                else
                    clearErrorType();
                // will be signed "corrected" in listBox
                refreshItemInListBox( m_number );
            }else{
                m_sFormula = sNewFormula;
                setPrecedentsCells();
                if( m_lNotValidPrecCells == null || m_lNotValidPrecCells.isEmpty()) {
                    clearErrorType();
                    refreshItemInListBox( m_number );
                }
                //  have to show the tpye becouse may be there are more than one kinds of faults in a formula
                //  in this case one formula cell can be seen more time in Validator with not same faults
                if( m_ErrorType == ERRORTYPE1 && getGui().m_errorType1 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE2 && getGui().m_errorType2 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE3 && getGui().m_errorType3 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE503 && getGui().m_errorType503 )
                    refreshItemInListBox( m_number );
/*
                if( m_ErrorType == ERRORTYPE5 && getGui().m_errorType5 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE6 && getGui().m_errorType6 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE7 && getGui().m_errorType7 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE8 && getGui().m_errorType8 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE9 && getGui().m_errorType9 )
                    refreshItemInListBox( m_number );
                if( m_ErrorType == ERRORTYPE10 && getGui().m_errorType10 )
                    refreshItemInListBox( m_number );
*/
            }
            
            if( _selectedItemPos == getNumber() ) {
                getGui().selectItemPosInListBox( m_number );
                getGui().showPrecedents();
                getGui().showDependents();
            }

            // need to adjust the selected item again
            if ( _selectedItemPos >= 0 )
                getGui().selectItemPosInListBox( _selectedItemPos );
            if ( _selectedItemPos < 0 && getGui().getListBoxSize() > 0 )
                getGui().selectItemPosInListBox( (short) 0 );
        }
    }

    public String getMemberCellSheetName(){
        String sheetName = "";
        if(m_number == _selectedItemPos ) {
            SingleReference cellRef = (SingleReference)getGui().getController().getRefToken( m_xCell ).Data;
            cellRef.Flags = 128;
            String activeCellName = getGui().getController().getActiveSheetName();
            String currSheetName = getGui().getController().getSheetNameByIndex(cellRef.Sheet);
            sheetName = currSheetName.equals(activeCellName) ? "" : currSheetName + ".";
        }
        return sheetName;
    }

    @Override
    public void disposing(EventObject arg0) {
        stop();
        clearPrecedentsCellsAndType();
        getGui().removeItemsFromLists();
    }

    @Override
    public synchronized void run() {
        try {
            while( m_isRunning ){
                setPrecedentsNewColor();
                Thread.sleep( 500 );
                setBackPrecedentsColors();
                Thread.sleep( 500 );
            }
        } catch (InterruptedException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void start(){
        if( m_ErrorType != 0 ){
            if( m_runner == null){
                m_runner = new Thread(this);
                m_isRunning = true;
                m_runner.start();
            }
        }
    }

    public void stop(){
        if (m_runner != null)
            m_runner = null;
        m_isRunning = false;
        setBackPrecedentsColors();
    }

}
