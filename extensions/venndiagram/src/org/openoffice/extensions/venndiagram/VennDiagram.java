/*****************************************************************
 * 
 * file: VennDiagram.java
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


import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.beans.PropertyValue;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.DispatchDescriptor;
import com.sun.star.frame.XController;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStatusListener;
import com.sun.star.lang.WrappedTargetException;
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
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;
import java.util.ArrayList;


public final class VennDiagram extends WeakBase
    implements XServiceInfo, XDispatchProvider, XInitialization, 
                                XDispatch, XSelectionChangeListener {
    
    private final        XComponentContext    m_xContext;
    private              XFrame               m_xFrame;
    private static final String               m_implementationName = VennDiagram.class.getName();
    private static final String[]             m_serviceNames = {"com.sun.star.frame.ProtocolHandler" };
    private              XController          m_xController;
    private              XSelectionSupplier   m_xSelectionSupplier;
    private              Controller           m_Controller;
    protected static     boolean              _bool = true;
    private static       ArrayList<XFrame>    frameList = null;
    
    public VennDiagram(XComponentContext context){
        m_xContext = context;
    };

    public static XSingleComponentFactory __getComponentFactory( String sImplementationName ) {
        XSingleComponentFactory xFactory = null;
        if(sImplementationName.equals(m_implementationName))
            xFactory = Factory.createComponentFactory(VennDiagram.class, m_serviceNames);
        return xFactory;
    }
    
    public static boolean __writeRegistryServiceInfo(XRegistryKey xRegistryKey) {
        return Factory.writeRegistryServiceInfo(m_implementationName, m_serviceNames, xRegistryKey);
    }

    // com.sun.star.lang.XServiceInfo:
    public String getImplementationName() {
         return m_implementationName;
    }

    public boolean supportsService(String sService){
        int len = m_serviceNames.length;
        for(int i=0;i<len;i++){
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
        if ( aURL.Protocol.compareTo("org.openoffice.extensions.venndiagram.venndiagram:") == 0 ){
            if(aURL.Path.compareTo("VennDiagram")== 0)
                return this;
        }
        return null;
    }

    // com.sun.star.frame.XDispatchProvider:
    public com.sun.star.frame.XDispatch[] queryDispatches(DispatchDescriptor[] seqDescriptors ){
        int nCount = seqDescriptors.length;
        XDispatch[] seqDispatcher = new XDispatch[seqDescriptors.length];
        for( int i=0; i < nCount; ++i ){
            seqDispatcher[i] = queryDispatch(seqDescriptors[i].FeatureURL,seqDescriptors[i].FrameName,seqDescriptors[i].SearchFlags );
        }
        return seqDispatcher;
    }

    // com.sun.star.lang.XInitialization:
    public void initialize(Object[] object)throws Exception{
        if (object.length>0){
            m_xFrame = (XFrame)UnoRuntime.queryInterface(XFrame.class, object[0]);
            m_xController = m_xFrame.getController();
            m_xSelectionSupplier = (XSelectionSupplier) UnoRuntime.queryInterface(XSelectionSupplier.class, m_xController);
            m_xSelectionSupplier.addSelectionChangeListener(this);
            m_Controller = new Controller(this, m_xContext, m_xFrame);             
            if(frameList == null){
                frameList = new ArrayList<XFrame>();
                frameList.add(m_xFrame);
            }else{
                boolean b = true;
                for(XFrame frame : frameList){
                    if(m_xFrame.equals(frame)){
                        b = false;
                    }
                }
                if(b){
                    frameList.add(m_xFrame);
                    _bool = true;
                }
            }
        }
    }
   
    // com.sun.star.frame.XDispatch:
     public void dispatch(URL aURL, PropertyValue[] aArguments){
         if ( aURL.Protocol.compareTo("org.openoffice.extensions.venndiagram.venndiagram:") == 0 ){
            if ( aURL.Path.compareTo("VennDiagram") == 0 ){
                m_Controller.createDiagram(3);
                return;    
            }
        }
    }
    public void addStatusListener(XStatusListener xControl, URL aURL ){ 
    
    }

    public void removeStatusListener(XStatusListener xControl, URL aURL){ 
    
    }
    
    public XShape getSelectedShape(){
        try {
            XShapes xShapes = (XShapes)UnoRuntime.queryInterface(XShapes.class, m_xSelectionSupplier.getSelection());
            if(xShapes != null)
                return (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(0));
        } catch (IndexOutOfBoundsException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } 
        return null;
    } 
     
    public void setSelectedShape(Object obj){
        try {
            m_xSelectionSupplier.select(obj);
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        }
    }
    
    //XSelectionChangeListener
    public void selectionChanged(EventObject event) {
        try {   
            if (m_Controller.getShapeName(getSelectedShape()).startsWith("VennDiagram")){
                m_Controller.setVisibleGui(true); 
                m_Controller.initDiagram();
            }
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (IndexOutOfBoundsException ex) {
            ex.printStackTrace();
        }catch (Exception ex) {
            ex.printStackTrace();
        } 
    }
    
    public void disposing(EventObject event) {
    }
    
} 
