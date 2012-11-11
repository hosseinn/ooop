package oxygenoffice.extensions.smart;

import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.frame.FrameActionEvent;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.frame.XStatusListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.ui.XContextMenuInterception;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.URL;
import java.util.ArrayList;


public final class SmartProtocolHandler extends WeakBase
   implements com.sun.star.lang.XInitialization,
              com.sun.star.frame.XDispatch,
              com.sun.star.lang.XServiceInfo,
              com.sun.star.frame.XDispatchProvider,
              XFrameActionListener
{
    private final        XComponentContext  m_xContext;
    private XFrame                          m_xFrame;
    private static final String             m_implementationName = SmartProtocolHandler.class.getName();
    private static final String[]           m_serviceNames = { "com.sun.star.frame.ProtocolHandler" };
    private              Controller         m_Controller     = null;
    private              XComponent         m_xComponent     = null;
    // store every frame with its Controller object
    private static ArrayList<FrameObject>   _frameObjectList  = null;
//    private              String             m_sOOName         = "";
//    private              String             m_sOOSetupVersion = "";


    public SmartProtocolHandler( XComponentContext context )
    {
        m_xContext = context;
//        setLONameAndVersion();
    };

/*
    public void setLONameAndVersion(){
        try {
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object oConfigurationProvider = xMCF.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XMultiServiceFactory xMSF = (XMultiServiceFactory)UnoRuntime.queryInterface(XMultiServiceFactory.class, oConfigurationProvider);
            String str[] = new String[1];
            str[0] = "/org.openoffice.Setup/Product";
            Object oConfigurationAccess = xMSF.createInstanceWithArguments("com.sun.star.configuration.ConfigurationAccess", str);
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, oConfigurationAccess);
            m_sOOName = AnyConverter.toString(xProps.getPropertyValue("ooName"));
            m_sOOSetupVersion = AnyConverter.toString(xProps.getPropertyValue("ooSetupVersion"));
        } catch (com.sun.star.uno.Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }    
    }
     
    public String getOOName(){ return m_sOOName; }
    public String getOOSetupVersion(){ return m_sOOSetupVersion; }
*/
    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;
        if ( sImplementationName.equals( m_implementationName ) )
            xFactory = Factory.createComponentFactory(SmartProtocolHandler.class, m_serviceNames);
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo( XRegistryKey xRegistryKey ) {
        return Factory.writeRegistryServiceInfo(m_implementationName, m_serviceNames, xRegistryKey);
    }

    // com.sun.star.lang.XInitialization:
    @Override
    public void initialize( Object[] object )
        throws com.sun.star.uno.Exception
    {
        if ( object.length > 0 )
        {
            m_xFrame = (com.sun.star.frame.XFrame)UnoRuntime.queryInterface(com.sun.star.frame.XFrame.class, object[0]);
            
            // add the m_xFrame and its m_Controller to the static arrayList of _frameObjectList
            // avoid the duplicate gui controls
            boolean isNewFrame = true;
            if(_frameObjectList == null){
                _frameObjectList = new ArrayList<FrameObject>();
            }else{
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                        isNewFrame = false;
            }
            if(isNewFrame){
                m_xFrame.addFrameActionListener(this);

                m_xComponent = (XComponent)UnoRuntime.queryInterface(XComponent.class, m_xFrame.getController().getModel());
                if(m_xComponent != null)
                    m_xComponent.addEventListener(this);

                if(m_Controller == null)
                    m_Controller = new Controller( this, m_xContext, m_xFrame );
                
                _frameObjectList.add(new FrameObject(m_xFrame, m_Controller));

                XContextMenuInterception xContextMenuInterception = (XContextMenuInterception)UnoRuntime.queryInterface(XContextMenuInterception.class, m_xFrame.getController());
                if(xContextMenuInterception != null)
                    xContextMenuInterception.registerContextMenuInterceptor(new ContextMenuInterceptor(m_xContext, m_xFrame, this));
            }else{
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                       m_Controller =  frameObj.getController();
            }
            threatDiagramExtensionIfItExists();
        }
    }

    public Controller getController(){
        return m_Controller;
    }

    public void threatDiagramExtensionIfItExists(){
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            String[][] str = xPIP.getExtensionList();
            for(int i = 0; i < str.length; i++)
                for(int j = 0; j < str[i].length; j++)
                    if(str[i][j].equals("org.openoffice.extensions.diagrams.Diagrams"))
                        new WarningThread(m_xContext, m_xFrame, xPIP).start();
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    // com.sun.star.frame.XDispatch:
    @Override
    public void dispatch( com.sun.star.util.URL aURL,
                           com.sun.star.beans.PropertyValue[] aArguments )
    {
         if ( aURL.Protocol.compareTo("oxygenoffice.extensions.smart.smartprotocolhandler:") == 0 )
        {
            if ( aURL.Path.compareTo("showSmartGallery") == 0 )
            {
                short exec = 0;
                if(m_Controller != null){
                    exec = m_Controller.executeGalleryDialog();
                    if(exec == 1)
                        m_Controller.createDiagram();
                }
                return;
            }
           
            if ( aURL.Path.compareTo("createDiagramFromList") == 0 )
            {
                if(m_Controller != null)
                    m_Controller.createDiagramFromList();
                return;
            }
/*
            if ( aURL.Path.compareTo("showOrHideRootElement") == 0 )
            {
                if(m_Controller != null)
                    m_Controller.showOrHideRootElement();
                return;
            }
*/
        }
    }

    @Override
    public void addStatusListener( XStatusListener xControl, URL aURL ) {
/*
        System.out.println(aURL.Complete);
        m_xStatusListener = xControl;
        m_createFromListURL = aURL;
        if (aURL.Complete.compareTo("oxygenoffice.extensions.smart.smartprotocolhandler:createDiagramFromList") == 0) {
            FeatureStateEvent aEvent = new FeatureStateEvent();
            aEvent.FeatureURL = aURL;
            aEvent.Source = (com.sun.star.frame.XDispatch) this;
            aEvent.IsEnabled = false;
            aEvent.Requery = false;
            xControl.statusChanged(aEvent);
        }
*/
    }

    @Override
    public void removeStatusListener( com.sun.star.frame.XStatusListener xControl,
                                       com.sun.star.util.URL aURL )
    {
        // add your own code here
    }

    // com.sun.star.lang.XServiceInfo:
    @Override
    public String getImplementationName() {
         return m_implementationName;
    }

    @Override
    public boolean supportsService( String sService ) {
        int len = m_serviceNames.length;
        for( int i=0; i < len; i++) {
            if (sService.equals(m_serviceNames[i]))
                return true;
        }
        return false;
    }

    @Override
    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.frame.XDispatchProvider:
    @Override
    public com.sun.star.frame.XDispatch queryDispatch(com.sun.star.util.URL aURL, String sTargetFrameName, int iSearchFlags)
    {
        if ( aURL.Protocol.compareTo("oxygenoffice.extensions.smart.smartprotocolhandler:") == 0 )
        {
            if ( aURL.Path.compareTo("showSmartGallery") == 0 )
                return this;

            if ( aURL.Path.compareTo("createDiagramFromList") == 0 )
                return this;
            
            if ( aURL.Path.compareTo("showOrHideRootElement") == 0 )
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    @Override
    public com.sun.star.frame.XDispatch[] queryDispatches(com.sun.star.frame.DispatchDescriptor[] seqDescriptors )
    {
        int nCount = seqDescriptors.length;
        com.sun.star.frame.XDispatch[] seqDispatcher = new com.sun.star.frame.XDispatch[seqDescriptors.length];
        for( int i=0; i < nCount; ++i )
        {
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL, seqDescriptors[i].FrameName, seqDescriptors[i].SearchFlags );
        }
        return seqDispatcher;
    }

     // XFrameActionListener
    @Override
    public void disposing(EventObject event) {
        // when the frame is closed we have to remove FrameObject item into the list
        if( event.Source.equals(m_xFrame)){
            m_xFrame.removeFrameActionListener(this);
            if(_frameObjectList != null){
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                        _frameObjectList.remove(frameObj);
            }
        }else{
            // when the document is closed we have to remove FrameObject item into the list
            m_xComponent.removeEventListener(this);
            m_Controller.getGui().closeAndDisposeControlDialog();
            if(_frameObjectList != null){
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                        _frameObjectList.remove(frameObj);
            }
        }
    }

    @Override
    public void frameAction(FrameActionEvent arg0) {
    }

}
