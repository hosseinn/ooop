
package org.openoffice.extensions.validator;

import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XCheckBox;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.document.XEventBroadcaster;
import com.sun.star.document.XEventListener;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XSheetAuditing;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.table.BorderLine;
import com.sun.star.table.CellAddress;
import com.sun.star.table.TableBorder;
import com.sun.star.table.XCell;
import com.sun.star.text.XText;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.XModifyBroadcaster;
import com.sun.star.util.XModifyListener;
import java.util.ArrayList;



public class Gui implements XModifyListener, XEventListener, XDialogEventHandler, XTopWindowListener {

    private     Controller              m_Controller            = null;
    private     XComponentContext       m_xContext              = null;
    private     XFrame                  m_xFrame                = null;
    private     XSpreadsheetDocument    m_xDocument             = null;
    private     XDialog                 m_xSelectDialog         = null;

    private     XCheckBox               m_xCheckBox1            = null;
    private     XCheckBox               m_xCheckBox2            = null;
    private     XCheckBox               m_xCheckBox3            = null;
    private     XCheckBox               m_xCheckBox4            = null;
/*
    private     XCheckBox               m_xCheckBox5            = null;
    private     XCheckBox               m_xCheckBox6            = null;
    private     XCheckBox               m_xCheckBox7            = null;
    private     XCheckBox               m_xCheckBox8            = null;
    private     XCheckBox               m_xCheckBox9            = null;
    private     XCheckBox               m_xCheckBox10           = null;
*/
    protected   boolean                 m_errorType1            = true; // empty cell references
    protected   boolean                 m_errorType2            = true; // string cell rererences
    protected   boolean                 m_errorType3            = true; // single references ( e.g.: 456)
    protected   boolean                 m_errorType503          = true; // division by zero
/*
    protected   boolean                 m_errorType5            = true;
    protected   boolean                 m_errorType6            = true;
    protected   boolean                 m_errorType7            = true;
    protected   boolean                 m_errorType8            = true;
    protected   boolean                 m_errorType9            = true;
    protected   boolean                 m_errorType10           = true;
*/
    private     XFixedText              m_xSheetFixedText       = null;
    private     XCheckBox               m_xPrecCheckBox         = null;
    private     XCheckBox               m_xDepCheckBox          = null;

    private     XDialog                 m_xDialog               = null;
    private     XWindow                 m_xWindow               = null;
    private     XTopWindow              m_xTopWindow            = null;
    private     XListBox                m_xListBox              = null;
    private     XFixedText              m_xFixedText            = null;
    private     XTextComponent          m_xTextComponent        = null;

    private     XEventBroadcaster       m_xDocEventBroadcaster  = null;
    private     XModifyBroadcaster      m_xDocModifyBroadcaster = null;
    private     int                     m_counter               = 0;

    private     short                   m_lastItemPos           = -1;
    private     short                   m_infoBoxOk             = -1;
    private     XPropertySet            m_currXPropSet          = null;
    private     Integer                 m_currCellColor         = null;
    private     TableBorder             m_sCellBorder           = null;
    private     boolean                 m_correctProcessRunning = false;


    Gui( Controller controller, XComponentContext xContext, XFrame xFrame, XSpreadsheetDocument document ) {

        m_Controller            = controller;
        m_xContext              = xContext;
        m_xFrame                = xFrame;
        m_xDocument             = document;
        m_xDocModifyBroadcaster = (XModifyBroadcaster) UnoRuntime.queryInterface(XModifyBroadcaster.class, m_xDocument);
        m_xDocModifyBroadcaster.addModifyListener(this);
        m_xDocEventBroadcaster = (XEventBroadcaster) UnoRuntime.queryInterface(XEventBroadcaster.class, m_xDocument);
        m_xDocEventBroadcaster.addEventListener(this);
        adjustCounter();
    }

    protected Controller getController(){
        return m_Controller;
    }

    public void setVisibleSelectWindow( boolean b ){
        if( m_xSelectDialog == null )
            createSelectDialog();
        if( b ) {
            if( m_xWindow != null )
               closeAndClearValidator();
            setErrorTypes( true );
            m_xSelectDialog.execute();
        } else {
            m_xSelectDialog.endExecute();
            m_xSelectDialog = null;
        }
    }


    public void createSelectDialog(){
        try {
            String sPackageURL              = getPackageLocation();
            String sDialogURL               = sPackageURL + "/dialogs/SelectWindow.xdl";
            XDialogProvider2 xDialogProv    = getDialogProvider();
            m_xSelectDialog                 = xDialogProv.createDialogWithHandler( sDialogURL, this );

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface( XControlContainer.class, m_xSelectDialog );

            Object oCheckBox    = xControlContainer.getControl("checkBox1");
            m_xCheckBox1        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox2");
            m_xCheckBox2        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox3");
            m_xCheckBox3        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox4");
            m_xCheckBox4        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
/*
            oCheckBox           = xControlContainer.getControl("checkBox5");
            m_xCheckBox5        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox6");
            m_xCheckBox6        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox7");
            m_xCheckBox7        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox8");
            m_xCheckBox8        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox9");
            m_xCheckBox9        = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
            oCheckBox           = xControlContainer.getControl("checkBox10");
            m_xCheckBox10       = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);
*/
         }catch(Exception ex){
            System.err.println("Exception in Gui.createDialog(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public void setVisibleValidator( boolean isVisible ){
        if( isVisible ){
            if( m_xWindow == null )
                createValidator();
            setErrorList();
        }
        m_xWindow.setVisible( isVisible );        
    }

    public void createValidator(){
        try {
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/Validator.xdl";
            XDialogProvider2 xDialogProv = getDialogProvider();
            m_xDialog = xDialogProv.createDialogWithHandler(sDialogURL, this);

            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xDialog);

            Object oListBox = xControlContainer.getControl("fixListBox");
            XListBox xFixListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class, oListBox);
            String cellHeader = createStringWithSpace(getDialogPropertyValue("Validator", "Validator.cellLabel.Label"), 22);
            String descHeader = getDialogPropertyValue("Validator", "Validator.errorLabel.Label");
            String label = cellHeader + descHeader;
            xFixListBox.addItem(label, (short) 0);

            Object oFixedText = xControlContainer.getControl("sheetNameLabel");
            m_xSheetFixedText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            oFixedText = xControlContainer.getControl("formulaLabel");
            m_xFixedText = (XFixedText) UnoRuntime.queryInterface(XFixedText.class, oFixedText);

            oListBox = xControlContainer.getControl("listBox");
            m_xListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class, oListBox);

            Object oTextField = xControlContainer.getControl("formulaField");
            m_xTextComponent = (XTextComponent) UnoRuntime.queryInterface(XTextComponent.class, oTextField);

            Object oCheckBox = xControlContainer.getControl("checkBox1");
            m_xPrecCheckBox = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);

            oCheckBox = xControlContainer.getControl("checkBox2");
            m_xDepCheckBox = (XCheckBox) UnoRuntime.queryInterface(XCheckBox.class, oCheckBox);


            if (m_xDialog != null) {
                m_xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, m_xDialog);
                //    This hasn't yet worked to make dialog window resizable.
                //    Short flags = (WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.SIZEABLE | WindowAttribute.CLOSEABLE);
                //    Rectangle posSize = xWindow.getPosSize()
                //    m_xWindow.setPosSize(posSize.X, posSize.Y, posSize.Width, posSize.Height, flags);
                m_xTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xWindow);
                m_xTopWindow.addTopWindowListener(this);
            }
        }catch(Exception ex){
            System.err.println("Exception in Gui.createDialog(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            location =  xPIP.getPackageLocation("org.openoffice.extensions.validator.Validator");
        } catch (NoSuchElementException ex) {
            System.err.println("Exception in Gui.getPackageLocation(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("Exception in Gui.getPackageLocation(). Message:\n" + ex.getLocalizedMessage());
        }
        return location;
    }

    public XDialogProvider2 getDialogProvider(){
        XDialogProvider2 xDialogProv = null;
        try {
            XModel xModel = m_xFrame.getController().getModel();
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object obj;
            if (xModel != null) {
                Object[] args = new Object[1];
                args[0] = xModel;
                obj = xMCF.createInstanceWithArgumentsAndContext("com.sun.star.awt.DialogProvider2", args, m_xContext);
            } else {
                obj = xMCF.createInstanceWithContext("com.sun.star.awt.DialogProvider2", m_xContext);
            }
            xDialogProv = (XDialogProvider2) UnoRuntime.queryInterface(XDialogProvider2.class, obj);
        }catch (Exception ex) {
            System.err.println("Exception in Gui.getDialogProvider(). Message:\n" + ex.getLocalizedMessage());
        }
        return xDialogProv;
    }

    public String getDialogPropertyValue( String sDialog, String name){
        String result = null;
        String m_resRootUrl = getPackageLocation() + "/dialogs/";
        XStringResourceWithLocation xResources = null;
        try {
            xResources = StringResourceWithLocation.create(m_xContext, m_resRootUrl, true, getController().getLocation(), sDialog, "", null);
        } catch (IllegalArgumentException ex) {
              ex.printStackTrace();
        }
        // map properties
        if(xResources != null){
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++) {
                if(ids[i].contains(name))
                    result = xResources.resolveString(ids[i]);
            }
        }
        return result;
    }

    public String createStringWithSpace(String s, int max){
        char c;
        int n;
        double dn = s.length()*2;
        for(int i=0; i<s.length();i++){
            c = s.charAt(i);
            if(c=='W')
                dn += 1.3333;
            if(c=='w'||c=='Q'||c=='O'||c=='Ó'||c=='Ö'||c=='Ő'||c=='m'||c=='M')
                dn += 0.6666;
            if(c=='A'||c=='R'||c=='U'||c=='Ú'||c=='Ü'||c=='Á'||c=='D'||c=='G'||c=='H'||c=='Ű'||c=='N'||c=='C')
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

    public void setErrorList() {
        getController().setNum((short)0);
        getController().setActiveSheet();
        m_xSheetFixedText.setText( getController().getActiveSheetName() );
        setTextComp("");
        m_lastItemPos = -1;

        if( m_errorType1 || m_errorType2 || m_errorType503 )
            getController().setErrorList();
        if( m_errorType3 )
            getController().setError3List();

        if( m_xListBox.getItemCount() > 0 ) {
            m_xListBox.selectItemPos( (short)0, true );
            m_lastItemPos = 0;
        }
    }

    public void setErrorTypes( boolean b ){
        short state = 0;
        if( b )
            state = 1;
        m_xCheckBox1.setState(state);
        m_xCheckBox2.setState(state);
        m_xCheckBox3.setState(state);
        m_xCheckBox4.setState(state);
/*
        m_xCheckBox5.setState(state);
        m_xCheckBox6.setState(state);
        m_xCheckBox7.setState(state);
        m_xCheckBox8.setState(state);
        m_xCheckBox9.setState(state);
        m_xCheckBox10.setState(state);
*/

        m_errorType1 = m_errorType2 = m_errorType3 = m_errorType503 = b;
/*
        m_errorType5 = m_errorType6 = m_errorType7 = m_errorType8 =
        m_errorType9 = m_errorType10 = b;
*/
    }

    public void refresh() {
        removeItemsFromLists();
        clearArrows();
        setErrorList();
    }

    public void removeItemsFromLists(){
        
        short itemPos = m_xListBox.getSelectedItemPos();
        if( itemPos >= 0 ){
             getController().getModel().getItemFromList( m_lastItemPos ).stop();
             setBackLastCellColor();
        }
        short n = -1;
        if( m_xListBox != null )
            n = m_xListBox.getItemCount();
        if(n > 0)
            m_xListBox.removeItems( (short)0, n );
        if( getController().getModel() != null )
            getController().getModel().clearList();
        m_lastItemPos = -1;
    }
    
    //listBox methods (5)
    public short getListBoxSize(){
        return m_xListBox.getItemCount();
    }

    public void addItemToListBox( String str, short n ){
        m_xListBox.addItem( str, n );
    }

    public void removeItemFromListBox(short itemPos){
        m_xListBox.removeItems(itemPos,(short)1);
    }

    public short getSelectedItemPosFromListBox(){
        return m_xListBox.getSelectedItemPos();
    }

    public void selectItemPosInListBox( short selectedItemPos ){
        m_xListBox.selectItemPos( selectedItemPos, true );
    }

    public void removeSelectionInListBox(){
        m_xListBox.selectItemPos( getSelectedItemPosFromListBox(), false );
    }

    public boolean isCorrectProcessRunning(){
        return m_correctProcessRunning;
    }

    @Override
    public boolean callHandlerMethod( XDialog xDialog, Object obj, String methodName ) throws WrappedTargetException {
        if(methodName.equals("signAll")){
            setErrorTypes( true );
            return true;
        }
        if(methodName.equals("clearAll")){
            setErrorTypes( false );
            return true;
        }
        if(methodName.equals("ok")){
            setVisibleSelectWindow( false );
            m_xSelectDialog = null;
            setVisibleValidator( true );
            return true;
        }
        if(methodName.equals("modifiedCheckBox1")){
            m_errorType1 = !m_errorType1;
            return true;
        }
        if(methodName.equals("modifiedCheckBox2")){
            m_errorType2 = !m_errorType2;
            return true;
        }
        if(methodName.equals("modifiedCheckBox3")){
            m_errorType3 = !m_errorType3;
            return true;
        }
        if(methodName.equals("modifiedCheckBox4")){
            m_errorType503 = !m_errorType503;
            return true;
        }
        /*
        if(methodName.equals("modifiedCheckBox5")){
            m_errorType5 = !m_errorType5;
            return true;
        }
        if(methodName.equals("modifiedCheckBox6")){
            m_errorType6 = !m_errorType6;
            return true;
        }
        if(methodName.equals("modifiedCheckBox7")){
            m_errorType7 = !m_errorType7;
            return true;
        }
        if(methodName.equals("modifiedCheckBox8")){
            m_errorType8 = !m_errorType8;
            return true;
        }
        if(methodName.equals("modifiedCheckBox9")){
            m_errorType9 = !m_errorType9;
            return true;
        }
        if(methodName.equals("modifiedCheckBox10")){
            m_errorType10 = !m_errorType10;
            return true;
        }
         *
         */
        if(methodName.equals("run")){
            refresh();
            return true;
        }
        if(methodName.equals("correct")){
            short itemPos = m_xListBox.getSelectedItemPos();
            Cell currCell = getController().getModel().getItemFromList(itemPos);
            if( currCell.getErrorType() > 0 ) {
                // m_correctProcessRunning is a semaphor to close
                // listening of Cell objects
                m_correctProcessRunning = true;
                if( currCell.getErrorType() == 3 )
                    correctErrorType3( currCell, itemPos );
                else
                    correctErrorType( currCell, itemPos );
                m_correctProcessRunning = false;
            }
            return true;
        }
        if(methodName.equals("showPrecedents")){
            showPrecedents();
            return true;
        }
        if(methodName.equals("showDependents")){
            showDependents();
            return true;
        }
        if(methodName.equals("itemChanged")){

            clearArrows();
            if ( m_xPrecCheckBox.getState() == 1 )
                showPrecedents(true);
            if ( m_xDepCheckBox.getState() == 1 )
                showDependents(true);
            
            short itemPos = m_xListBox.getSelectedItemPos();
            if ( m_lastItemPos >= 0 ) {
                Cell lastCell = getController().getModel().getItemFromList( m_lastItemPos );
                lastCell.stop();
            }
            setBackLastCellColor();

            Cell currCell = getController().getModel().getItemFromList(itemPos);

            XCell xCurrCell = currCell.getXCell();
            if( currCell.getErrorType() == 3 ) {
                currCell.setPrecedentsCells();
                m_xFixedText.setText( getDialogPropertyValue("Strings", "Strings.valueLabel.Label" ) );
                XText xText = (XText)UnoRuntime.queryInterface( XText.class, xCurrCell );
                setTextComp( xText.getString() );
            } else {
                setCurrCellNewColor( xCurrCell ); // it is yellow except in case of errorType3
                m_xFixedText.setText( getDialogPropertyValue("Validator", "Validator.formulaLabel.Label") );
                setTextComp( getController().getFormula(xCurrCell) );
            }
            if(currCell.getErrorType() >= 0)
                currCell.start();
         
            m_lastItemPos = itemPos;

            return true;
        }
        if(methodName.equals("taskManager")){
            closeAndClearValidator();
            setVisibleSelectWindow( true );
            return true;
        }
        return false;
    }

    public void setBackLastCellColor(){
        if ( m_lastItemPos >= 0 && m_currXPropSet != null && m_currCellColor != null ) {
            try {
                m_currXPropSet.setPropertyValue("CellBackColor", new Integer(m_currCellColor));
                m_currXPropSet.setPropertyValue("TableBorder", m_sCellBorder);
            } catch (Exception ex) {
                System.err.println("Exception in Gui.windowClosing(). Message:\n" + ex.getLocalizedMessage());
            }
        }
    }

    public void setCurrCellNewColor( XCell xCell ){
        try {
            m_currXPropSet = (com.sun.star.beans.XPropertySet) UnoRuntime.queryInterface( com.sun.star.beans.XPropertySet.class, xCell );
            m_currCellColor = (Integer) m_currXPropSet.getPropertyValue( "CellBackColor" );
            m_sCellBorder = (TableBorder) m_currXPropSet.getPropertyValue( "TableBorder" );

            m_currXPropSet.setPropertyValue("CellBackColor", new Integer(0xFFFF00));
            m_currXPropSet.setPropertyValue("TableBorder", getTableBorder(new Integer(0x000000), (short)100 ));
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in Gui.setCurrCellNewColor(). Message:\n" + ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in Gui.setCurrCellNewColor(). Message:\n" + ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in Gui.setCurrCellNewColor(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in Gui.setCurrCellNewColor(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public TableBorder getTableBorder(Integer color, short width){
        BorderLine borderLine = new BorderLine();
        borderLine.Color = color;
        borderLine.OuterLineWidth = width;
        TableBorder tableBorder = new TableBorder();
        tableBorder.VerticalLine = tableBorder.HorizontalLine = tableBorder.LeftLine =
        tableBorder.RightLine = tableBorder.TopLine = tableBorder.BottomLine =
        borderLine;
        tableBorder.IsVerticalLineValid = tableBorder.IsHorizontalLineValid = tableBorder.IsLeftLineValid =
        tableBorder.IsRightLineValid = tableBorder.IsTopLineValid = tableBorder.IsBottomLineValid =
        true;
        return tableBorder;
    }

    public void correctErrorType( Cell currCell, short itemPos ) {

        String newFormula = m_xTextComponent.getText();
        ArrayList<XCell> newFormulaPrecedents = null;
        if( !newFormula.equals("") )
            newFormulaPrecedents = getController().getPrecedentCellsOfFormula( newFormula );

        if( !newFormula.equals("") && ( newFormulaPrecedents == null || newFormulaPrecedents.isEmpty() ) ) {
            String sFormula = getController().getFormula( currCell.getXCell());
            showMessageBox( getDialogPropertyValue("Strings", "Strings.errorWindowTitle.Label" ), newFormula + " - " + getDialogPropertyValue("Strings", "Strings.invalidFormula.Label" ) + " " +sFormula);
            setTextComp( sFormula );
        } else {
            XCell currXCell = currCell.getXCell();

            if( newFormula.equals("") ) {
                currCell.stop();
                currCell.clearErrorType();
                clearArrows();
                currXCell.setFormula( newFormula );
                currCell.setFormula( getController().getFormula( currXCell ) );
                currCell.refreshItemInListBoxAndSelect(itemPos);
            } else {
                String oldFormula = getController().getFormula( currXCell );

                short sValid = getController().parseXCells( newFormulaPrecedents );
                if( sValid == 0 ) {
                    currXCell.setFormula( newFormula );
                    sValid = (short)currXCell.getError();
                }
                if( sValid > 0 ) {
                    currXCell.setFormula( oldFormula );
                    if( sValid == 532 )
                        sValid = 503;
                    String message = getDialogPropertyValue("Strings", "Strings.errorMessage" + sValid + ".Label" );
                    if( message == null )
                        message = "error:" + sValid;
                    showMessageBox( getDialogPropertyValue("Strings", "Strings.errorWindowTitle.Label" ), newFormula + " - " + message );
                    setTextComp( getController().getFormula( currXCell ) );
                }else{
                    currCell.stop();
                    clearArrows();
                    currCell.removeListener();
                    currCell.setFormula( getController().getFormula(currXCell) );
                    currCell.setPrecedentsCells();
                    currCell.clearErrorType();
                    currCell.refreshItemInListBoxAndSelect(itemPos);
                    setTextComp( getController().getFormula(currXCell) );
                }
            }
        }
    }
/*
    public void correctErrorType( Cell currCell, short itemPos ) {
        String normalFormula = currCell.getXCell().getFormula();
        System.out.println("normalFormula : " + normalFormula);
        String oldFormula = getController().getFormula( currCell.getXCell() );
        System.out.println("oldFormula : " + oldFormula);
        String oldUSFormula = getController().conversSpecificFNameToOriginalFName(oldFormula);
        System.out.println("oldUSFormula : " + oldUSFormula);
        String newFormula = m_xTextComponent.getText(); // newFormula can be special languages formula
        System.out.println("newFormula : " + newFormula);
        String newUSFormula = getController().conversSpecificFNameToOriginalFName(newFormula);
        System.out.println("newUSFormula : " + newUSFormula);
        ArrayList<XCell> newFormulaPrecedents = null;
        if( !newFormula.equals("") )
            newFormulaPrecedents = getController().getPrecedentCellsOfFormula( newFormula );
       
        if( !newFormula.equals("") && ( newFormulaPrecedents == null || newFormulaPrecedents.isEmpty() ) ) {
            showMessageBox( getDialogPropertyValue("Strings", "Strings.errorWindowTitle.Label" ), newFormula + " - " + getDialogPropertyValue("Strings", "Strings.invalidFormula.Label" ) + " " + oldFormula);
            setTextComp( oldFormula );
        }else {
            XCell currXCell = currCell.getXCell();
            if( newFormula.equals("") ) {
                currCell.stop();
                currCell.clearErrorType();
                clearArrows();
                currXCell.setFormula( newFormula );
                currCell.setFormula( getController().getFormula( currXCell ) );
                currCell.refreshItemInListBoxAndSelect(itemPos);
            }else {
                //oldFormula = getController().conversSpecificFNameToOriginalFName(oldFormula);
                short sValid = getController().parseXCells( newFormulaPrecedents );
                // if( sValid == 0 || sValid == 2 ) { allow correct strings arguments to strings also
                if( sValid == 0 ) {
                    currXCell.setFormula( newUSFormula );
                    sValid = (short)currXCell.getError();
                }
                //if( sValid == 0 || sValid > 2) {
                if( sValid > 0 ) {
                    currXCell.setFormula( oldUSFormula );
                    if( sValid == 532 )
                        sValid = 503;
                    String message = getDialogPropertyValue("Strings", "Strings.errorMessage" + sValid + ".Label" );
                    if( message == null )
                        message = "error:" + sValid;
                    showMessageBox( getDialogPropertyValue("Strings", "Strings.errorWindowTitle.Label" ), newFormula + " - " + message );
                    setTextComp( getController().getFormula( currXCell ) );
                }else{
                    currCell.stop();
                    clearArrows();
                    currCell.removeListener();
                    currCell.setFormula( getController().getFormula(currXCell) );
                    currCell.setPrecedentsCells();
                    currCell.clearErrorType();
                    currCell.refreshItemInListBoxAndSelect(itemPos);
                    setTextComp( getController().getFormula(currXCell) );
                }
            }
        }
    }
*/
    public void correctErrorType3( Cell currCell, short itemPos ) {

        String newValue = m_xTextComponent.getText();
        if( !newValue.equals( "" ) )
            showMessageBox2( getDialogPropertyValue("Strings", "Strings.errorMessage3.Label" ), getDialogPropertyValue("Strings", "Strings.infoBoxMessage.Label" ) );
        if( newValue.equals( "" ) || m_infoBoxOk == 1 ) {
            currCell.stop();
            currCell.clearPrecedentsCellsAndType();
            XCell currXCell = currCell.getXCell();
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, currXCell);
            xText.setString("");
            currCell.refreshItemInListBoxAndSelect(itemPos);
            setTextComp("");
        }
        m_infoBoxOk = -1;
    }

    public void setTextComp(String str){
        m_xTextComponent.setText( str );
    }

    public void showPrecedents(){
        if ( m_xPrecCheckBox.getState() == 1 )
            showPrecedents(true);
        else
            showPrecedents(false);
    }

    public void showPrecedents( boolean bool ){
        CellAddress addr = getListBoxSelectedItemAddress();
        XSheetAuditing xSheetAuditing = (XSheetAuditing)UnoRuntime.queryInterface( XSheetAuditing.class, getController().getActiveSheet() );
        if( bool )
            xSheetAuditing.showPrecedents(addr);
        else
            xSheetAuditing.hidePrecedents(addr);
    }

    public void showDependents(){
        if ( m_xDepCheckBox.getState() == 1 )
            showDependents(true);
        else
            showDependents(false);
    }

    public void showDependents( boolean bool ){
        CellAddress addr = getListBoxSelectedItemAddress();
        XSheetAuditing xSheetAuditing = (XSheetAuditing)UnoRuntime.queryInterface( XSheetAuditing.class, getController().getActiveSheet() );
        if( bool )
            xSheetAuditing.showDependents(addr);
        else
            xSheetAuditing.hideDependents(addr);
    }

    public void clearArrows(){
        XSheetAuditing xSheetAuditing = null;
        xSheetAuditing = (XSheetAuditing) UnoRuntime.queryInterface( XSheetAuditing.class, getController().getActiveSheet() );
        if(xSheetAuditing != null)
            xSheetAuditing.clearArrows();
    }

    public void clearArrowsOfClosingSheet(){
        XSheetAuditing xSheetAuditing = null;
        xSheetAuditing = (XSheetAuditing) UnoRuntime.queryInterface( XSheetAuditing.class, getController().getClosingActiveSheet() );
        if(xSheetAuditing != null)
            xSheetAuditing.clearArrows();
    }

    public CellAddress getListBoxSelectedItemAddress(){
        short itemPos = m_xListBox.getSelectedItemPos();
        XCell xCell = getController().getModel().getItemFromList( itemPos ).getXCell();
        XCellAddressable xCellAddr =  (XCellAddressable)UnoRuntime.queryInterface(XCellAddressable.class, xCell);
        return xCellAddr.getCellAddress();
    }

    public void showMessageBox(String sTitle, String sMessage){
        try{
            Object oToolkit = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", m_xContext);
            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);
            if ( m_xFrame != null && xToolkit != null ) {
                WindowDescriptor aDescriptor = new WindowDescriptor();
                aDescriptor.Type              = WindowClass.MODALTOP;
                aDescriptor.WindowServiceName = new String( "infobox" );
                aDescriptor.ParentIndex       = -1;
                aDescriptor.Parent            = (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class, m_xFrame.getContainerWindow());
                aDescriptor.Bounds            = new Rectangle(0,0,300,200);
                aDescriptor.WindowAttributes  = WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.CLOSEABLE;
                XWindowPeer xPeer = xToolkit.createWindow( aDescriptor );
                if ( null != xPeer ) {
                   XMessageBox xMessageBox = (XMessageBox)UnoRuntime.queryInterface(XMessageBox.class, xPeer);
                    if ( null != xMessageBox ){
                        xMessageBox.setCaptionText( sTitle );
                        xMessageBox.setMessageText( sMessage );
                        m_xWindow.setEnable(false);
                        xMessageBox.execute();
                        m_xWindow.setEnable(true);
                        m_xWindow.setFocus();
                    }
                }
            }
        } catch (Exception ex) {
            System.err.println("Exception in Gui.showMessageBox(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public void showMessageBox2(String sTitle, String sMessage){
        XComponent xComponent = null;
        try {
            Object oToolkit = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit", m_xContext);
            XToolkit xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);
            XMessageBoxFactory xMessageBoxFactory = (XMessageBoxFactory)UnoRuntime.queryInterface(XMessageBoxFactory.class, oToolkit);
            if ( m_xFrame != null && xToolkit != null ) {
                WindowDescriptor aDescriptor = new WindowDescriptor();
                aDescriptor.Type              = WindowClass.MODALTOP;
                aDescriptor.ParentIndex       = -1;
                aDescriptor.Parent            = (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class, m_xFrame.getContainerWindow());
                XWindowPeer xPeer = xToolkit.createWindow( aDescriptor );
                if ( null != xPeer ) {
                    // rectangle may be empty if position is in the center of the parent peer
                    Rectangle aRectangle = new Rectangle();
                    XMessageBox xMessageBox = xMessageBoxFactory.createMessageBox( xPeer, aRectangle, "messbox", com.sun.star.awt.MessageBoxButtons.BUTTONS_OK_CANCEL, sTitle, sMessage);
                    xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                    if (xMessageBox != null){
                        m_xWindow.setEnable(false);
                        m_infoBoxOk = xMessageBox.execute();
                        m_xWindow.setEnable(true);
                        m_xWindow.setFocus();
                    }
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
             System.err.println("Exception in Gui.showMessageBox2(). Message:\n" + ex.getLocalizedMessage());
        }finally{
            if (xComponent != null)
                xComponent.dispose();
        }
    }

    @Override
    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[19];
        aMethods[0] = "signAll";
        aMethods[1] = "clearAll";
        aMethods[2] = "ok";
        aMethods[3] = "modifiedCheckBox1";
        aMethods[4] = "modifiedCheckBox2";
        aMethods[5] = "modifiedCheckBox3";
        aMethods[6] = "modifiedCheckBox4";
        aMethods[7] = "modifiedCheckBox5";
        aMethods[8] = "modifiedCheckBox6";
        aMethods[9] = "modifiedCheckBox7";
        aMethods[10] = "modifiedCheckBox8";
        aMethods[11] = "modifiedCheckBox9";
        aMethods[12] = "modifiedCheckBox10";
        aMethods[13] = "run";
        aMethods[14] = "correct";
        aMethods[15] = "showPrecedents";
        aMethods[16] = "showDependents";
        aMethods[17] = "itemChanged";
        aMethods[18]= "taskManager";
        return aMethods;
    }

    public void adjustCounter(){
        m_counter = m_xDocument.getSheets().getElementNames().length;
    }

    public int countSheets(){
        return m_xDocument.getSheets().getElementNames().length;
    }

    // XModifyListener (2)
    // listen exsisting spreadseets (user can delete or make a new spreadseet)
    // if the structure of document is modulated, the list will be refreshed
    @Override
    public void modified(EventObject event) {
        if( m_counter != countSheets() )  {
            adjustCounter();
            refresh();
        }
    }

    @Override
    public void disposing(EventObject event) {
        m_xDocModifyBroadcaster.removeModifyListener(this);
        closeAndClearValidator();
        if( m_xSelectDialog != null )
            m_xSelectDialog = null;
    }

    public void closeAndClearValidator(){
        removeItemsFromLists();
        clearArrowsOfClosingSheet();
        if(m_xWindow != null){
            m_xWindow.setVisible(false);
            m_xWindow = null;
        }
        if(m_xDialog != null)
            m_xDialog = null;
    }

    // XTopWindowListener (7)
    @Override
    public void windowClosing(EventObject event) {
        closeAndClearValidator();
        if( m_xSelectDialog != null )
            m_xSelectDialog = null;
    }

    @Override
    public void windowClosed(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void windowOpened(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void windowMinimized(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void windowNormalized(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void windowActivated(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void windowDeactivated(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    @Override
    public void notifyEvent(com.sun.star.document.EventObject event) {
        if(event.EventName.equals("OnPrepareViewClosing")){
            getController().getModel().stop();
            setBackLastCellColor();
            removeSelectionInListBox();
        }
        if(event.EventName.equals("OnViewClosed"))
            m_xDocEventBroadcaster.removeEventListener(this); 
    }
}