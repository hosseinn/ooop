package org.openoffice.extensions.watchwindow;

import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.DispatchDescriptor;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStatusListener;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XInitialization;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.util.URL;


public final class WatchWindow extends WeakBase 
                 implements XServiceInfo, XDispatchProvider, XInitialization, XDispatch {

    private final XComponentContext m_xContext;
    private com.sun.star.frame.XFrame m_xFrame;
    private Controller m_Controller = null;
    private static final String m_implementationName = WatchWindow.class.getName();
    private static final String[] m_serviceNames = { "com.sun.star.frame.ProtocolHandler" };

    public WatchWindow( XComponentContext context ){
        m_xContext = context;
    };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ){
        XSingleComponentFactory xFactory = null;
        if (sImplementationName.equals(m_implementationName))
            xFactory = Factory.createComponentFactory(WatchWindow.class, m_serviceNames);
        return xFactory;
    }

    public static boolean __writeRegistryServiceInfo( XRegistryKey xRegistryKey ){
        return Factory.writeRegistryServiceInfo(m_implementationName,m_serviceNames,xRegistryKey);
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
    public com.sun.star.frame.XDispatch queryDispatch(URL aURL, String sTargetFrameName, int iSearchFlags){
        if (aURL.Protocol.compareTo("org.openoffice.extensions.watchwindow.watchwindow:")== 0){
            if(aURL.Path.compareTo("Display")== 0)
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public XDispatch[] queryDispatches(DispatchDescriptor[] seqDescriptors ){
        int nCount = seqDescriptors.length;
        XDispatch[] seqDispatcher = new XDispatch[seqDescriptors.length];
        for(int i=0;i<nCount;++i){
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,
                                             seqDescriptors[i].FrameName,
                                             seqDescriptors[i].SearchFlags );
        }
        return seqDispatcher;
    }

    // com.sun.star.lang.XInitialization:
    public void initialize(Object[] object) throws com.sun.star.uno.Exception{
        if (object.length>0){
            m_xFrame = (XFrame)UnoRuntime.queryInterface(XFrame.class, object[0]);
        }
    }
   
    // com.sun.star.frame.XDispatch:
     public void dispatch(URL aURL,PropertyValue[] aArguments ){
         if(aURL.Protocol.compareTo("org.openoffice.extensions.watchwindow.watchwindow:")== 0){
            if(aURL.Path.compareTo("Display")== 0){
                try {
                    if (m_Controller == null) {
                        m_Controller = new Controller(m_xContext, m_xFrame);
                    }
                    m_Controller.displayWatchWindow();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
                return;
            }
        }
    }
    
    public void addStatusListener(XStatusListener xControl,URL aURL) {
        
    }

    public void removeStatusListener(XStatusListener xControl,URL aURL) {
    
    }
  
    public void disposing(EventObject arg0) {
        throw new UnsupportedOperationException("Not supported yet.");
    }
    
} 