package oxygenoffice.extensions.smart;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexContainer;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.ui.ActionTriggerSeparatorType;
import com.sun.star.ui.ContextMenuExecuteEvent;
import com.sun.star.ui.ContextMenuInterceptorAction;
import com.sun.star.ui.XContextMenuInterceptor;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class ContextMenuInterceptor implements XContextMenuInterceptor {

    XComponentContext       m_xContext  = null;
    XFrame                  m_xFrame    = null;
    SmartProtocolHandler    m_oSmartPH  = null;

    ContextMenuInterceptor(XComponentContext xContext, XFrame xFrame, SmartProtocolHandler oSmartPH) {
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_oSmartPH  = oSmartPH;
    }

    @Override
    public ContextMenuInterceptorAction notifyContextMenuExecute(ContextMenuExecuteEvent aEvent) {
        try {
            Object oSelection = aEvent.Selection.getSelection();
            XShape xSelectedShape = m_oSmartPH.getController().getSelectedShape();
            XIndexContainer xContextMenu = aEvent.ActionTriggerContainer;
            XPropertySet xFirsElement = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xContextMenu.getByIndex(0));
            String sFirstElementName = AnyConverter.toString(xFirsElement.getPropertyValue("CommandURL"));
            if(isShapeCollectionService(oSelection) && isShapeService(xSelectedShape) && !sFirstElementName.equals("slot:27014")){
                XMultiServiceFactory xMenuElementFactory = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, xContextMenu);
                if (xMenuElementFactory != null) {
                    XPropertySet xRootMenuEntry = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                    XPropertySet xSeparator = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator"));
                    Short aSeparatorType = new Short(ActionTriggerSeparatorType.LINE);
                    xSeparator.setPropertyValue("SeparatorType", (Object)aSeparatorType );
                    xContextMenu.insertByIndex ( xContextMenu.getCount(), (Object)xSeparator );
                    xContextMenu.insertByIndex ( xContextMenu.getCount(), (Object)xRootMenuEntry );
                    String commandText = m_oSmartPH.getController().getGui().getDialogPropertyValue("Strings2", "Strings2.CreateFromList.Label");
                    xRootMenuEntry.setPropertyValue("Text", commandText);
                    xRootMenuEntry.setPropertyValue("CommandURL", "oxygenoffice.extensions.smart.smartprotocolhandler:createDiagramFromList");
                    xRootMenuEntry.setPropertyValue("Image", getGraphic("/images/smart_16.png"));
/*                    
                    if(xSelectedShape != null){
                        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, xSelectedShape);
                        if(xNamed != null && m_oSmartPH.getController().isSmARTOrganigramShape(xNamed.getName())){
                            if(((OrganizationChart)m_oSmartPH.getController().getDiagram()).getDiagramTree().getRootItem().isFirstChild()){
                                XPropertySet xRootMenuEntry2 = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                                xContextMenu.insertByIndex ( xContextMenu.getCount(), (Object)xRootMenuEntry2 );
                                String commandText2 = "";
                                if(m_oSmartPH.getController().isHiddenRootElement())
                                    commandText2 = m_oSmartPH.getController().getGui().getDialogPropertyValue("Strings2", "Strings2.ShowRootElement.Label");
                                else
                                    commandText2 = m_oSmartPH.getController().getGui().getDialogPropertyValue("Strings2", "Strings2.HideRootElement.Label");
                                xRootMenuEntry2.setPropertyValue("Text", commandText2);
                                xRootMenuEntry2.setPropertyValue("CommandURL", "oxygenoffice.extensions.smart.smartprotocolhandler:showOrHideRootElement");
                                xRootMenuEntry2.setPropertyValue("Image", getGraphic("/images/smart_16.png"));
                            }
                        }
                    }
*/
                    return com.sun.star.ui.ContextMenuInterceptorAction.EXECUTE_MODIFIED ;
                }
            }
        } catch ( com.sun.star.beans.UnknownPropertyException ex ) {
            System.err.println(ex.getLocalizedMessage());
        } catch ( com.sun.star.lang.IndexOutOfBoundsException ex ) {
          System.err.println(ex.getLocalizedMessage());
        } catch ( com.sun.star.uno.Exception ex ) {
          System.err.println(ex.getLocalizedMessage());
        } catch ( java.lang.Throwable ex ) {
          System.err.println(ex.getLocalizedMessage());
        }
        return com.sun.star.ui.ContextMenuInterceptorAction.IGNORED;
    }

    public boolean isShapeCollectionService(Object obj){
        if(obj != null){
            //m_oSmartPH.getController().testServices(obj);
            XServiceInfo xSI = (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, obj);
            String[] str = xSI.getSupportedServiceNames();
            for(String s : str)
                if(s.equals("com.sun.star.drawing.ShapeCollection"))
                    return true;
        }
        return false;
    }

    public boolean isShapeService(Object obj){
        boolean isShape = false;
        if(obj != null){
            XServiceInfo xSI = (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, obj);
            String[] str = xSI.getSupportedServiceNames();
            for(String s : str){
                if(s.equals("com.sun.star.drawing.Shape")) //com.sun.star.drawing.TextShape
                    isShape = true;
                if(s.equals("com.sun.star.drawing.GroupShape"))
                    return false;
            }
        }
        return isShape;
    }

    public XGraphic getGraphic(String sImageUrl){
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
            return xGraphicProvider.queryGraphic(aPropertyValues);
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
        return null;
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

}
