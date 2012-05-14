package com.example;

import com.sun.star.frame.FrameActionEvent;
import com.sun.star.frame.XFrameActionListener;
import com.sun.star.lang.EventObject;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.lib.uno.helper.Factory;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.registry.XRegistryKey;
import com.sun.star.lib.uno.helper.WeakBase;


public final class StarMaker extends WeakBase
   implements com.sun.star.lang.XInitialization,
              com.sun.star.frame.XDispatch,
              com.sun.star.lang.XServiceInfo,
              com.sun.star.frame.XDispatchProvider,
              XFrameActionListener
{
    private final XComponentContext         m_xContext;
    private com.sun.star.frame.XFrame       m_xFrame;
    private static final String             m_implementationName = StarMaker.class.getName();
    private static final String[]           m_serviceNames = { "com.sun.star.frame.ProtocolHandler" };

    private DrawStar                        m_DrawStar      = null;

    public StarMaker( XComponentContext context )
    {
        m_xContext = context;
    };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;
        if ( sImplementationName.equals( m_implementationName ) )
            xFactory = Factory.createComponentFactory(StarMaker.class, m_serviceNames);
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
        }
    }

    // com.sun.star.frame.XDispatch:
    @Override
    public void dispatch( com.sun.star.util.URL aURL,
                           com.sun.star.beans.PropertyValue[] aArguments )
    {
        if ( aURL.Protocol.compareTo("com.example.starmaker:") == 0 )
        {
            if ( aURL.Path.compareTo("StarMaker") == 0 )
            {
                // add your code here
                if(m_DrawStar == null)
                    m_DrawStar = new DrawStar(m_xContext, m_xFrame);
                m_DrawStar.showStarMakerDialog();
                return;
            }
        }
    }

    @Override
    public void addStatusListener( com.sun.star.frame.XStatusListener xControl,
                                    com.sun.star.util.URL aURL )
    {
        // add your own code here
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
        if ( aURL.Protocol.compareTo("com.example.starmaker:") == 0 )
        {
            if ( aURL.Path.compareTo("StarMaker") == 0 )
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

    }

    @Override
    public void frameAction(FrameActionEvent arg0) {

    }
}
