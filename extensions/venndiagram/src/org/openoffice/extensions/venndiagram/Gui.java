/*****************************************************************
 * 
 * file: Gui.java
 * 
 * The Venn Diagram Extension has been licensed under term of
 * GNU Lesser General Public License (LGPL). 
 * You can read and understand LGPL license here:
 * http://www.gnu.org/licenses/lgpl.html
 * 
 * This extension is created by: Tibor Hornyák and
 * University of Szeged, Hungary - http://www.u-szeged.hu/english/
 * OxygenOffice Professional Team - http://ooop.sf.net/
 * 
 * license_en-US document V1.0.0
 * 
 ****************************************************************/

package org.openoffice.extensions.venndiagram;

import com.sun.star.awt.ItemEvent;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogEventHandler;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XTopWindow;
import com.sun.star.awt.XTopWindowListener;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class Gui implements XTopWindowListener, XDialogEventHandler {
    
    private Controller             m_Controller        = null;
    private XComponentContext      m_xContext          = null; 
    private XModel                 m_xModel            = null;
    private XMultiComponentFactory m_xServiceManager   = null;
    private XDialog                m_xDialogPanel      = null;
    private XWindow                m_xWindowPanel      = null;
    private XTopWindow             m_xTopWindow        = null;
    private XTopWindow             m_xPaletteTopWindow = null;
    private boolean                m_status            = false;
    private XControl               m_xImageControl     = null;
    private XDialog                m_xPaletteDialog    = null;

   
    Gui(Controller controller, XComponentContext xContext, XModel xModel) {
        try {
            m_Controller = controller;
            m_xContext = xContext;
            m_xModel = xModel;
            m_xServiceManager = m_xContext.getServiceManager();
            createControlPanel();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (NoSuchElementException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    public void setVisible(boolean bool){
        if(!m_status && bool || m_status && !bool){
            m_xWindowPanel.setVisible(bool);
            m_status = bool;   
        }
    }
    
    public void createControlPanel() throws IllegalArgumentException, NoSuchElementException, WrappedTargetException, Exception {
        XDialogProvider2 xDialogProv = getDialogProvider();              
        String sPackageURL = getPackageLocation();
        String sDialogURL = sPackageURL + "/dialogs/GuiPanel.xdl";
        m_xDialogPanel = xDialogProv.createDialogWithHandler( sDialogURL, this );
        if(m_xDialogPanel != null){
            m_xWindowPanel = (XWindow)UnoRuntime.queryInterface(XWindow.class, m_xDialogPanel);
            m_xTopWindow = (XTopWindow)UnoRuntime.queryInterface(XTopWindow.class, m_xWindowPanel); 
            m_xTopWindow.addTopWindowListener(this); 
        } 
        XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xDialogPanel);       
        m_xImageControl = xControlContainer.getControl("ImageControl");  
        setBackgroundColor(0x00FF00);
        setVisible(true);
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
            location =  xPIP.getPackageLocation("org.openoffice.extensions.venndiagram.VennDiagram");
        } catch (NoSuchElementException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        }
        return location;
    }
    
    public void showColorTable() throws IllegalArgumentException, NoSuchElementException, WrappedTargetException, Exception{
        XDialogProvider2 xDialogProv = getDialogProvider();              
        String sPackageURL = getPackageLocation();     
        String sDialogURL = sPackageURL + "/dialogs/ColorTable.xdl";
        m_xPaletteDialog = xDialogProv.createDialogWithHandler( sDialogURL, this );
        if(m_xPaletteDialog != null){
            m_xPaletteTopWindow = (XTopWindow)UnoRuntime.queryInterface(XTopWindow.class, m_xPaletteDialog);
            m_xPaletteTopWindow.addTopWindowListener(this);
            m_xWindowPanel.setEnable(false);
            m_xPaletteDialog.execute();
         }
    }
    
    public void setColor(XDialog xDialog, int i){
        try {
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, xDialog);
            XControl xImageControl = xControlContainer.getControl("ImageControl" + i);
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xImageControl.getModel());
            int color = AnyConverter.toInt(xPropImage.getPropertyValue("BackgroundColor"));
            setBackgroundColor(color);
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (UnknownPropertyException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        }
    }
    
    public void setBackgroundColor(int color){
        try {
            XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xImageControl.getModel());
            xPropImage.setPropertyValue("BackgroundColor", new Integer(color));
        } catch (UnknownPropertyException ex) {
            ex.printStackTrace();
        } catch (PropertyVetoException ex) {
            ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        }
        
    }
    
    public int getBackgroundColor(){
         XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xImageControl.getModel());
         int color = 0;
        try {
            color = AnyConverter.toInt(xPropImage.getPropertyValue("BackgroundColor"));
        } catch (UnknownPropertyException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        }
         return color;
    }
    
    public int getNum(String name){
        String s ="";
        char[] charName = name.toCharArray();
        int i = 5;
        while(i<name.length())
           s +=  charName[i++];
        return getController().parseInt(s);
    }
     
    //XDialogEventHandler (2)
    public String[] getSupportedMethodNames() {
        String[] aMethods = new String[52];
        aMethods[0] = "addShape";
        aMethods[1] = "removeShape";
        aMethods[2] = "changeMode";
        aMethods[3] = "showColorTable";
        String methodName = "";
        for(int i=1;i<=48;i++){
            methodName = "image" +i;
            aMethods[i+3] = methodName;
        }
        return aMethods; 
    }
    
    public boolean callHandlerMethod(XDialog xDialog, Object EventObject, String MethodName){
        try {
            //addShape
            if(MethodName.equals("addShape")){
                if( getController().getDrawPage() != null && getController().getShapes() != null){
                    getController().addShape();
                    getController().refreshDiagram();
                    return true;
                }
            }
            //removeShape
            if(MethodName.equals("removeShape")){ 
                if( getController().getDrawPage() != null && getController().getShapes() != null){
                    getController().removeShape();
                    getController().refreshDiagram();
                    return true;
                }
            }
            //changeMode
            if(MethodName.equals("changeMode")){
                getController().changeMode((short)((ItemEvent)EventObject).Selected);
                if( getController().getDrawPage() != null && getController().getShapes() != null){
                    getController().refreshShapeProperties();
                    getController().refreshDiagram();
                    return true;
                }
            }
            //show ColorTable
            if(MethodName.equals("showColorTable")){ 
                showColorTable();
                return true;
            }
            //user define the color
            if(MethodName.contains("image")){ 
                int num = getNum(MethodName); 
                setColor(xDialog, num);
                m_xPaletteDialog.endExecute();
                m_xWindowPanel.setEnable(true);
                return true;
            }
        } catch (IndexOutOfBoundsException ex) {
           ex.printStackTrace();
        } catch (UnknownPropertyException ex) {
           ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
           ex.printStackTrace();
        } catch (Exception ex) {
           ex.printStackTrace();
        }
        return false;
    }
    
    //XTopWindowListener (8)
    public void windowClosing(EventObject event) {
        if(event.Source.equals(m_xWindowPanel))
            setVisible(false);
        if(event.Source.equals(m_xPaletteTopWindow)){
           m_xPaletteTopWindow.removeTopWindowListener(this);
           m_xPaletteDialog.endExecute();
           m_xWindowPanel.setEnable(true);
        }
    }
    
    public void windowOpened(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
    public void windowClosed(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
    public void windowMinimized(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
    public void windowNormalized(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    public void windowActivated(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
    public void windowDeactivated(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
    public void disposing(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
} 
