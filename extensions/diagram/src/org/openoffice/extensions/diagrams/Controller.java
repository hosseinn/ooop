package org.openoffice.extensions.diagrams;

import com.sun.star.container.XNamed;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;
import org.openoffice.extensions.diagrams.diagram.Diagram;
import org.openoffice.extensions.diagrams.diagram.cyclediagram.CycleDiagram;
import org.openoffice.extensions.diagrams.diagram.organizationdiagram.OrganizationDiagram;
import org.openoffice.extensions.diagrams.diagram.pyramiddiagram.PyramidDiagram;
import org.openoffice.extensions.diagrams.diagram.venndiagram.VennDiagram;


public class Controller implements XSelectionChangeListener {

    private         XComponentContext   m_xContext              = null;
    private         XFrame              m_xFrame                = null;
    private         XController         m_xController           = null;
    private         Gui                 m_Gui                   = null;
    private         XSelectionSupplier  m_xSelectionSupplier    = null;

    private         String              m_sLastDiagramName      = "";

    private         Diagram             m_Diagram               = null;
    private         short               m_DiagramType;
    private         short               m_LastDiagramType       = 0;
    private         int                 m_LastDiagramID         = -1;

    protected static final short        ORGANIGRAM              = 1;
    protected static final short        VENNDIAGRAM             = 2;
    protected static final short        PYRAMIDDIAGRAM          = 3;
    protected static final short        CYCLEDIAGRAM            = 4;


    Controller( XComponentContext xContext, XFrame xFrame ){
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_xController = m_xFrame.getController();
        setGui();
        addSelectionListener();
    }

    public Diagram getDiagram(){
        return m_Diagram;
    }

    public void removeDiagram(){
        if(m_Diagram != null)
            m_Diagram = null;
    }

    public void setLastDiagramName(String name){
        m_sLastDiagramName = name;
    }

    public void setDiagramType(short dType){
        m_DiagramType = dType;
    }

    public short getDiagramType(){
        return m_DiagramType;
    }

    public void setLastDiagramType(short dType){
        m_LastDiagramType = dType;
    }

    public short getLastDiagramType(){
        return m_LastDiagramType;
    }

    public void setLastDiagramID(int iD){
        m_LastDiagramID = iD;
    }

    public int getLastDiagramID(){
        return m_LastDiagramID;
    }

    public void addSelectionListener(){
        if(m_xSelectionSupplier == null)
            m_xSelectionSupplier = (XSelectionSupplier) UnoRuntime.queryInterface(XSelectionSupplier.class, m_xController);
        if(m_xSelectionSupplier != null)
            m_xSelectionSupplier.addSelectionChangeListener(this);
    }

    public void removeSelectionListener(){
        if(m_xSelectionSupplier != null)
            m_xSelectionSupplier.removeSelectionChangeListener(this);
    }

    public Gui getGui(){
        return m_Gui;
    }

    public void setGui(){
        if(m_Gui == null){
            if(m_xContext != null && m_xFrame != null)
                m_Gui = new Gui(this, m_xContext, m_xFrame);
        }
    }

    public void setVisibleSelectWindow(boolean bool){
        if(m_Gui != null)
            m_Gui.setVisibleSelectWindow(true);
    }

    public XDrawPage getCurrentPage(){
        XDrawView xDrawView = (XDrawView)UnoRuntime.queryInterface(XDrawView.class, m_xController);
        return xDrawView.getCurrentPage();
    }

    public Locale getLocation() {
        Locale locale = null;
        try {
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object oConfigurationProvider = xMCF.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XLocalizable xLocalizable = (XLocalizable) UnoRuntime.queryInterface(XLocalizable.class, oConfigurationProvider);
            locale = xLocalizable.getLocale();
        } catch (Exception ex) {
            System.err.println("Exception in Gui.getLocation(). Message:\n" + ex.getLocalizedMessage());
        }
        return locale;
    }

    // adjust the DiagramId during init()
    public String getCurrentDiagramIdName() {
        String name = getDiagram().getShapeName(getSelectedShape());
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length() &&  charName[i] != '-')
           s +=  charName[i++];
        return s;
    }

    public String getNumberStrOfShape(String name){
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  charName[i] != '-')
            i++;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length())
           s +=  charName[i++];
        return s;
    }

    public int getNumberOfShape(String name){
        return parseInt(getNumberStrOfShape(name));
    }

    public int parseInt(String s) {
        int n = -1;
        try{
            n = Integer.parseInt(s);
        }catch(NumberFormatException ex){
             ex.printStackTrace();
        }
        return n;
    }

    // XSelectionChangeListener
    @Override
    public void disposing(EventObject arg0) {
        
    }

 
    // XSelectionChangeListener
    @Override
    public void selectionChanged(EventObject event) {
        XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, getSelectedShape() );
        String selectedShapeName = xNamed.getName();
        // listen the diagrams
        if(selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("VennDiagram") || selectedShapeName.startsWith("PyramidDiagram") || selectedShapeName.startsWith("CycleDiagram")) {
            String newDiagramName = selectedShapeName.split("-", 2)[0];
            // if the previous selected item not in same diagram,
            // need to instantiate the new diagram
            if( m_sLastDiagramName.equals("") || !m_sLastDiagramName.equals(newDiagramName)){
                if( selectedShapeName.startsWith("OrganizationDiagram") )
                    setDiagramType(ORGANIGRAM);
                if( selectedShapeName.startsWith("VennDiagram") )
                    setDiagramType(VENNDIAGRAM);
                if( selectedShapeName.startsWith("PyramidDiagram") )
                    setDiagramType(PYRAMIDDIAGRAM);
                if( selectedShapeName.startsWith("CycleDiagram") )
                    setDiagramType(CYCLEDIAGRAM);
                instantiateDiagram();
                m_sLastDiagramName = newDiagramName;
                getGui().setVisibleControlDialog(true);
                getDiagram().initDiagram();
            }
            if(selectedShapeName.startsWith("OrganizationDiagram") && selectedShapeName.endsWith("RectangleShape0"))
                if(getDiagram() != null)
                    ((OrganizationDiagram)getDiagram()).selectGroupShape();
            //always need to init becouse the user can modify the page???
            if(!getGui().isVisibleControlDialog())
                getGui().setVisibleControlDialog(true);
            if(getGui().isVisibleControlDialog())
                getGui().setFocusControlDialog();
        }
    }

    public void instantiateDiagram(){
        if(m_DiagramType == ORGANIGRAM)
            m_Diagram = new OrganizationDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == VENNDIAGRAM)
            m_Diagram = new VennDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == PYRAMIDDIAGRAM)
            m_Diagram = new PyramidDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == CYCLEDIAGRAM)
            m_Diagram = new CycleDiagram(this, getGui(), m_xFrame);
    }

    public void setSelectedShape(Object obj){
        try {
           m_xSelectionSupplier.select(obj);
        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in Controller.setSelectedShape(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public XShapes getSelectedShapes(){
        return (XShapes) UnoRuntime.queryInterface(XShapes.class, m_xSelectionSupplier.getSelection());
    }

    public XShape getSelectedShape(){
        try {
            XShapes xShapes = getSelectedShapes();
            if (xShapes != null)
                return (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(0));
        } catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in Controller.getSelectedShape(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in Controller.getSelectedShape(). Message:\n" + ex.getLocalizedMessage());
        }
        return null;
    }

}
