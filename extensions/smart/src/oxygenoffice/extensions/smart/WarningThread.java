package oxygenoffice.extensions.smart;

import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowClass;
import com.sun.star.awt.WindowDescriptor;
import com.sun.star.awt.XMessageBox;
import com.sun.star.awt.XMessageBoxFactory;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XExtensionManager;
import com.sun.star.deployment.XPackage;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.deployment.XPackageManager;
import com.sun.star.deployment.XPackageManagerFactory;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.task.XAbortChannel;
import com.sun.star.ucb.XCommandEnvironment;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


public class WarningThread extends Thread{

    private XComponentContext           m_xContext  = null;
    private XFrame                      m_xFrame    = null;
    private XPackageInformationProvider m_xPIP      = null;
    private short                       m_sMessageValue;

    public WarningThread(XComponentContext xContext, XFrame xFrame, XPackageInformationProvider xPIP){
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_xPIP      = xPIP;
        m_sMessageValue = -1;
    }

    @Override
    public void run() {
        String title = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashWarning.Title.Label");
        String message = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashWarning.Message.Label");
        showWarningBox(title, message);
        if(m_sMessageValue == 1){
            if(removeDiagramsExtension()){
                title = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashDone.Title.Label");
                message = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashDone.Message.Label");
                showMessageBox(title, message);
            } else {
                title = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashHasnotDone.Title.Label");
                message = getDialogPropertyValue("Strings2", "Strings2.ExtensionCrashHasnotDone.Message.Label");
                showWarningBox(title, message);
            }
        }
    }

   

    public String getDialogPropertyValue(String dialogName, String propertyName){
        String result = null;
        XStringResourceWithLocation xResources = null;
        String resRootUrl = m_xPIP.getPackageLocation("oxygenoffice.extensions.smart.SmART") + "/dialogs/";
        try {
            xResources = StringResourceWithLocation.create(m_xContext, resRootUrl, true, getLocation(), dialogName, "", null);
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        if(xResources != null){
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++)
                if(ids[i].contains(propertyName))
                    result = xResources.resolveString(ids[i]);
        }
        return result;
    }

    public Locale getLocation() {
        Locale locale = null;
        try {
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object oConfigurationProvider = xMCF.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XLocalizable xLocalizable = (XLocalizable) UnoRuntime.queryInterface(XLocalizable.class, oConfigurationProvider);
            locale = xLocalizable.getLocale();
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return locale;
    }

    public void showWarningBox(String sTitle, String sMessage){
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
                    Rectangle aRectangle = new Rectangle();
                    XMessageBox xMessageBox = xMessageBoxFactory.createMessageBox( xPeer, aRectangle, "warningbox", com.sun.star.awt.MessageBoxButtons.BUTTONS_OK_CANCEL, sTitle, sMessage);
                    if (xMessageBox != null){
                        m_sMessageValue = xMessageBox.execute();
                        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                        if(xComponent != null)
                            xComponent.dispose();
                    }
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
             System.err.println(ex.getLocalizedMessage());
        }
    }

    public void showMessageBox(String sTitle, String sMessage){
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
                    Rectangle aRectangle = new Rectangle();
                    XMessageBox xMessageBox = xMessageBoxFactory.createMessageBox( xPeer, aRectangle, "messbox", com.sun.star.awt.MessageBoxButtons.BUTTONS_OK, sTitle, sMessage);
                    if (xMessageBox != null){
                        xMessageBox.execute();
                        XComponent xComponent = (XComponent) UnoRuntime.queryInterface(XComponent.class, xMessageBox);
                        if(xComponent != null)
                            xComponent.dispose();
                    }
                }
            }
        } catch (com.sun.star.uno.Exception ex) {
             System.err.println(ex.getLocalizedMessage());
        }
    }

    public boolean removeDiagramsExtension(){
        try{
            String location =  m_xPIP.getPackageLocation("org.openoffice.extensions.diagrams.Diagrams");
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            XExtensionManager xEM = (XExtensionManager) UnoRuntime.queryInterface(XExtensionManager.class, xNameAccess.getByName("/singletons/com.sun.star.deployment.ExtensionManager"));

            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            XCommandEnvironment xCE = (XCommandEnvironment) UnoRuntime.queryInterface(XCommandEnvironment.class, xMCF.createInstanceWithContext("com.sun.star.ucb.CommandEnvironment", m_xContext));

            XPackageManagerFactory xPMF = (XPackageManagerFactory) UnoRuntime.queryInterface(XPackageManagerFactory.class, xNameAccess.getByName("/singletons/com.sun.star.deployment.thePackageManagerFactory"));
            XPackageManager xPM = null;

            XPackage xPackage = xEM.getDeployedExtension("user", "org.openoffice.extensions.diagrams.Diagrams", location, xCE);
            if(xPackage != null){
                xPM = xPMF.getPackageManager("user");
            }else{
                xPackage = xEM.getDeployedExtension("share", "org.openoffice.extensions.diagrams.Diagrams", location, xCE);
                if(xPackage != null)
                    xPM = xPMF.getPackageManager("share");
            }

            XAbortChannel xACh = xPM.createAbortChannel();
            //xEM.disableExtension(xPackage, xACh, xCE);
            xEM.removeExtension("org.openoffice.extensions.diagrams.Diagrams", location, "user", xACh, xCE);

            boolean isDiagramExtenisonExist = false;
            String[][] str = m_xPIP.getExtensionList();
            for(int i = 0; i < str.length; i++)
                for(int j = 0; j < str[i].length; j++)
                    if(str[i][j].equals("org.openoffice.extensions.diagrams.Diagrams"))
                        isDiagramExtenisonExist = true;
            return !isDiagramExtenisonExist;
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
            return false;
        }
    }

}
