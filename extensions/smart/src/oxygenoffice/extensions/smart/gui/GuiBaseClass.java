package oxygenoffice.extensions.smart.gui;

import com.sun.star.awt.ImageAlign;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import oxygenoffice.extensions.smart.Controller;


public class GuiBaseClass {

    protected     XComponentContext       m_xContext                  = null;
    protected     XFrame                  m_xFrame                    = null;
    protected     Controller              m_Controller                = null;
    
    public GuiBaseClass() { }

    public GuiBaseClass(Controller controller, XComponentContext xContext, XFrame xFrame){
        m_Controller = controller;
        m_xContext = xContext;
        m_xFrame = xFrame;
    }

    public Controller getController(){
        return m_Controller;
    }

    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            location =  xPIP.getPackageLocation("oxygenoffice.extensions.smart.SmART");
        } catch (NoSuchElementException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
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
            System.err.println(ex.getLocalizedMessage());
        }
        return xDialogProv;
    }

    public String getDialogPropertyValue(String dialogName, String propertyName){
        String result = null;
        XStringResourceWithLocation xResources = null;
        String m_resRootUrl = getPackageLocation() + "/dialogs/";
        try {
            xResources = StringResourceWithLocation.create(m_xContext, m_resRootUrl, true, getController().getLocation(), dialogName, "", null);
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        // map properties
        if(xResources != null){
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++)
                if(ids[i].contains(propertyName))
                    result = xResources.resolveString(ids[i]);
        }
        return result;
    }

    

    public XGraphic setGraphic(XControl xImageControl, String sImageUrl){
        XGraphic xGraphic = null;
        try{
            Object oGraphicProvider = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.graphic.GraphicProvider", m_xContext);
            XGraphicProvider xGraphicProvider = (XGraphicProvider) UnoRuntime.queryInterface(XGraphicProvider.class, oGraphicProvider);
            PropertyValue[] aPropertyValues = new PropertyValue[1];
            PropertyValue aPropertyValue = new PropertyValue();
            aPropertyValue.Name = "URL";
            if(!sImageUrl.equals(""))
                aPropertyValue.Value = getPackageLocation() + sImageUrl;
            else
                aPropertyValue.Value = "";
            aPropertyValues[0] = aPropertyValue;
            xGraphic = xGraphicProvider.queryGraphic(aPropertyValues);
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xImageControl.getModel());
            xProps.setPropertyValue("Graphic", xGraphic);
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xGraphic;
    }

    public void setControlScaleImageProp(XControl xControl, boolean bool){
        try{
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            xProps.setPropertyValue("ScaleImage", new Boolean(bool));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setControlScaleModeProp(XControl xControl, short scaleMode){
        try{
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
            xProps.setPropertyValue("ScaleMode", new Short(scaleMode));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void enableControl(XControl xControl, boolean bool){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("Enabled", new Boolean(bool));
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

    public void visibleControl(XControl xControl, boolean bool){
        //XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, obj);
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("Visible", new Boolean(bool));
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

    public void enableVisibleControl(XControl xControl, boolean bool){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("EnableVisible", bool);
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

    public void setImageColorOfControl(XControl xControl, int color){
        if(xControl != null){
            try {
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                xPropImage.setPropertyValue("BackgroundColor", new Integer(color));
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public int getImageColorOfControl(XControl xControl){
        int color = -1;
        try {
            if(xControl != null){
                XPropertySet xPropImage = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                color = AnyConverter.toInt(xPropImage.getPropertyValue("BackgroundColor"));
            }
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return color;
    }

    public void setImageOfObject(Object object, String sImageUrl, short imageAlign){
        if(object != null){
            XControl xControl = (XControl) UnoRuntime.queryInterface(XControl.class, object);
            if(xControl != null){
                try {
                    XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xControl.getModel());
                    if(imageAlign == ImageAlign.LEFT || imageAlign == ImageAlign.RIGHT || imageAlign == ImageAlign.TOP || imageAlign == ImageAlign.BOTTOM)
                        xProps.setPropertyValue("ImageAlign", new Short(imageAlign));
                    xProps.setPropertyValue("ImageURL", sImageUrl);
                } catch (PropertyVetoException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (IllegalArgumentException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (UnknownPropertyException ex) {
                    System.err.println(ex.getLocalizedMessage());
                }
            }
        }
    }

    public int getNum(String name){
        String s ="";
        char[] charName = name.toCharArray();
        int i = 5;
        while(i<name.length())
           s +=  charName[i++];
        return getController().parseInt(s);
    }

}
