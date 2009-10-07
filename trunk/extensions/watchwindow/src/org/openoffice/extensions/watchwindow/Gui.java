package org.openoffice.extensions.watchwindow;

import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowAttribute;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XListBox;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.beans.PropertyValue;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.sheet.RangeSelectionEvent;
import com.sun.star.sheet.XRangeSelection;
import com.sun.star.sheet.XRangeSelectionListener;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class Gui implements XTopWindowListener, XDialogEventHandler, XRangeSelectionListener{
    
    private Controller             m_Controller        = null;
    private XComponentContext      m_xContext          = null;
    private XFrame                 m_xFrame            = null;
    private XModel                 m_xModel            = null;
    private XMultiComponentFactory m_xServiceManager   = null;
    private XToolkit               m_xToolkit          = null;
    private XWindow                m_xWindow           = null;
    private XTopWindow             m_xTopWindow        = null;
    private XRangeSelection        m_xRangeSelection   = null;
    private XListBox               m_xFixListBox       = null;
    private XListBox               m_xListBox          = null;
    private String                 m_sHandlerMethod1   = "addCell";
    private String                 m_sHandlerMethod2   = "removeCell";
    private String                 m_selectAreaName    = ""; 

    Gui(Controller controller, XComponentContext xContext, XModel xModel) throws Exception {
       m_Controller = controller;
       m_xContext = xContext;
       m_xModel = xModel;
       m_xFrame = m_xModel.getCurrentController().getFrame();
       m_xServiceManager = m_xContext.getServiceManager();
       Object oToolkit = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.awt.Toolkit",m_xContext);
       m_xToolkit = (XToolkit) UnoRuntime.queryInterface(XToolkit.class, oToolkit);       
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    void setVisible(boolean bool){
        if (m_xWindow == null){
            createWatchWindow();
        }
        m_xWindow.setVisible(bool);  
    }
    
    public void createWatchWindow(){
        try {
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/WWindow.xdl";
            XDialogProvider2 xDialogProv = getDialogProvider();
            XDialog xDialog = xDialogProv.createDialogWithHandler(sDialogURL, this);
            if (xDialog != null) {
                m_xWindow = (XWindow) UnoRuntime.queryInterface(XWindow.class, xDialog);
                m_xTopWindow = (XTopWindow) UnoRuntime.queryInterface(XTopWindow.class, m_xWindow);
                m_xTopWindow.addTopWindowListener(this);
            }
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, xDialog);
            Object oListBox = xControlContainer.getControl("FixListBox");
            m_xFixListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class, oListBox);
            String sheet = getController().createStringWithSpace(getDialogPropertyValue("WWindow.Sheet.Label"), 26);
            String cell = getController().createStringWithSpace(getDialogPropertyValue("WWindow.Cell.Label"), 14);
            String value = getController().createStringWithSpace(getDialogPropertyValue("WWindow.Value.Label"), 25);
            String formula = getDialogPropertyValue("WWindow.Formula.Label");
            String label = sheet + cell + value + formula;
            m_xFixListBox.addItem(label, (short) 0);
            oListBox = xControlContainer.getControl("ListBox1");
            m_xListBox = (XListBox) UnoRuntime.queryInterface(XListBox.class, oListBox);
            
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    
    public void addToListBox(String s, short n){
        m_xListBox.addItem( s, n );
    }
    
    public void removeFromListBox(short num, short n){
        m_xListBox.removeItems( num, n);
    }
    
    public short getSelectedListBoxPos(){
        return m_xListBox.getSelectedItemPos();
    }
    
    public XDialogProvider2 getDialogProvider(){
        XDialogProvider2 xDialogProv = null;
        try {
            Object obj;
            if (m_xModel != null) {
                Object[] args = new Object[1];
                args[0] = m_xModel;
                obj = m_xServiceManager.createInstanceWithArgumentsAndContext("com.sun.star.awt.DialogProvider2", args, m_xContext);
            } else {
    
                obj = m_xServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", m_xContext);
            } 
            xDialogProv = (XDialogProvider2) UnoRuntime.queryInterface(XDialogProvider2.class, obj);
        }catch (Exception ex) {
             ex.printStackTrace();
        }  
        return xDialogProv; 
    }

    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            location =  xPIP.getPackageLocation("org.openoffice.extensions.watchwindow.WatchWindow");
        } catch (NoSuchElementException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        }
        return location;
    }
    
    public String getDialogPropertyValue(String name){
        String result = null;
        try {
            String m_resRootUrl = getPackageLocation() + "/dialogs/";
            XStringResourceWithLocation xResources = StringResourceWithLocation.create(m_xContext, m_resRootUrl, true, getController().getLocation(), "WWindow", "", null);
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++) {
                if(ids[i].contains(name))
                    result = xResources.resolveString(ids[i]);
            }
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        }
        return result;
    }
    
    public void showMessageBox(int num) {
        String title = getDialogPropertyValue("WWindow.Title" +num +".Label");
        String message = getDialogPropertyValue("WWindow.Message" +num +".Label");    
        showMessageBox(title, message); 
    }
    
    public void showMessageBox(String sTitle, String sMessage){
        try{
         if ( m_xFrame != null && m_xToolkit != null ) {
                WindowDescriptor aDescriptor = new WindowDescriptor();
                aDescriptor.Type              = WindowClass.MODALTOP;
                aDescriptor.WindowServiceName = new String( "infobox" );
                aDescriptor.ParentIndex       = -1;
                aDescriptor.Parent            = (XWindowPeer)UnoRuntime.queryInterface(XWindowPeer.class, m_xFrame.getContainerWindow());
                aDescriptor.Bounds            = new Rectangle(0,0,300,200);
                aDescriptor.WindowAttributes  = WindowAttribute.BORDER | WindowAttribute.MOVEABLE | WindowAttribute.CLOSEABLE;   
                XWindowPeer xPeer = m_xToolkit.createWindow( aDescriptor );
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
            ex.printStackTrace();
        }
    }
    
    public void cellSelection(String cellName) {
        m_xRangeSelection = (XRangeSelection)UnoRuntime.queryInterface(XRangeSelection.class, m_xModel.getCurrentController());
        m_xRangeSelection.addRangeSelectionListener(this); 
        String title = getController().getLocation().Language.equals("hu")?"Cella kiválasztása":"Select cell";
        PropertyValue[] args = new PropertyValue[3];
        args[0] = new PropertyValue();
        args[0].Name = "Title";
        args[0].Value = title;
        args[1] = new PropertyValue();
        args[1].Name = "SingleCellMode";
        args[1].Value = true;
        args[2] = new PropertyValue();
        args[2].Name = "InitialValue";
        args[2].Value = cellName;
        setVisible(false);
        m_xRangeSelection.startRangeSelection(args);  
    }
    
    //com.sun.star.awt.XDialogEventHandler: (2)
    public boolean callHandlerMethod(XDialog arg0, Object arg1, String MethodName) {
        if(MethodName.equals(m_sHandlerMethod1)){
            getController().addCell();
            return true;
        }
        if(MethodName.equals(m_sHandlerMethod2)){
            getController().removeCell();
            return true;
        }
        return false;
    }

    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[2];
        aMethods[0] = m_sHandlerMethod1;
        aMethods[1] = m_sHandlerMethod2;
        return aMethods; 
    }
    
    //com.sun.star.sheet.XRangeSelectionListener: (2)
    public void done(RangeSelectionEvent event) {
         try {
            m_selectAreaName = event.RangeDescriptor;
            m_xRangeSelection.removeRangeSelectionListener(this);
            getController().done(m_selectAreaName);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void aborted(RangeSelectionEvent event) {
        m_xRangeSelection.removeRangeSelectionListener(this);
        setVisible(true);
    }
    
    //com.sun.star.awt.XTopWindowListener: (8)
    public void windowClosing(EventObject event) {
        m_xWindow.setVisible(false);
    }

    public void windowOpened(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void windowClosed(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void windowMinimized(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void windowNormalized(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void windowActivated(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void windowDeactivated(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }

    public void disposing(EventObject event) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
} 