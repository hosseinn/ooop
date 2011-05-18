package oxygenoffice.extensions.smart;

import com.sun.star.frame.FrameActionEvent;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import java.util.ArrayList;


public final class SmartProtocolHandler extends WeakBase
   implements com.sun.star.lang.XInitialization,
              com.sun.star.frame.XDispatch,
              com.sun.star.lang.XServiceInfo,
              com.sun.star.frame.XDispatchProvider,
              XFrameActionListener
{
    private final XComponentContext         m_xContext;
    private com.sun.star.frame.XFrame       m_xFrame;
    private static final String             m_implementationName = SmartProtocolHandler.class.getName();
    private static final String[]           m_serviceNames = { "com.sun.star.frame.ProtocolHandler" };
    private Controller                      m_Controller = null;
    private XComponent                      m_xComponent = null;
    // store every frame with its Controller object
    private static ArrayList<FrameObject>   _frameObjectList = null;


    public SmartProtocolHandler( XComponentContext context )
    {
        m_xContext = context;
    };

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
                
            }else{
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                       m_Controller =  frameObj.getController();
            }
        }
    }

    // com.sun.star.frame.XDispatch:
     public void dispatch( com.sun.star.util.URL aURL,
                           com.sun.star.beans.PropertyValue[] aArguments )
    {
         if ( aURL.Protocol.compareTo("oxygenoffice.extensions.smart.smartprotocolhandler:") == 0 )
        {
            if ( aURL.Path.compareTo("showSmartGallery") == 0 )
            {
                if(m_Controller != null)
                    m_Controller.executeGalleryDialog();
                return;
            }
        }
    }

    public void addStatusListener( com.sun.star.frame.XStatusListener xControl,
                                    com.sun.star.util.URL aURL )
    {
        // add your own code here
    }

    public void removeStatusListener( com.sun.star.frame.XStatusListener xControl,
                                       com.sun.star.util.URL aURL )
    {
        // add your own code here
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
         return m_implementationName;
    }

    public boolean supportsService( String sService ) {
        int len = m_serviceNames.length;
        for( int i=0; i < len; i++) {
            if (sService.equals(m_serviceNames[i]))
                return true;
        }
        return false;
    }

    public String[] getSupportedServiceNames() {
        return m_serviceNames;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch queryDispatch(com.sun.star.util.URL aURL, String sTargetFrameName, int iSearchFlags)
    {
        if ( aURL.Protocol.compareTo("oxygenoffice.extensions.smart.smartprotocolhandler:") == 0 )
        {
            if ( aURL.Path.compareTo("showSmartGallery") == 0 )
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
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
        // when the document is closed we have to remove FrameObject item into the list
        if(event.Source.toString().contains("com.sun.star.frame.XModel")){
            m_xComponent.removeEventListener(this);
            m_Controller.getGui().closeAndDisposeControlDialog();
            if(_frameObjectList != null){
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                        _frameObjectList.remove(frameObj);
            }
        }
        // when the frame is closed we have to remove FrameObject item into the list
        if( event.Source.equals(m_xFrame)){
            m_xFrame.removeFrameActionListener(this);
            if(_frameObjectList != null){
                for(FrameObject frameObj : _frameObjectList)
                    if(m_xFrame.equals(frameObj.getXFrame()))
                        _frameObjectList.remove(frameObj);
            }
        }
    }

    @Override
    public void frameAction(FrameActionEvent arg0) { }

}
