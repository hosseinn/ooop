package oxygenoffice.extensions.smart;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
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
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XText;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.horizontalorgchart.HorizontalOrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.simpleorgchart.SimpleOrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.tablehierarchyorgchart.TableHierarchyOrgChart;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.ShapeData;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.buttdiagram.TargetDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram.CycleDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram.PyramidDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram.VennDiagram;



public final class Controller implements XSelectionChangeListener {

    private         SmartProtocolHandler    m_SmartPH                   = null;
    private         XComponentContext       m_xContext                  = null;
    private         XFrame                  m_xFrame                    = null;
    private         XController             m_xController               = null;
    private         Gui                     m_Gui                       = null;
    private         XSelectionSupplier      m_xSelectionSupplier        = null;

    private         String                  m_sLastDiagramName          = "";

    private         Diagram                 m_Diagram                   = null;
    private         short                   m_DiagramType;
    private         short                   m_GroupType;
    private         short                   m_LastDiagramType           = -1;
    private         int                     m_LastDiagramID             = -1;
    private         int                     m_NumberOfPages             = 0;

    private         boolean                 m_IsOrganiGroupConversAction= false;

    public static final short               ORGANIGROUP                 = 0;
    public static final short               RELATIONGROUP               = 1;
    public static final short               LISTGROUP                   = 2;
    public static final short               PROCESSGROUP                = 3;
    public static final short               MATRIXGROUP                 = 4;
    
    public static final short               NOTDIAGRAM                  = -1;
    public static final short               SIMPLEORGANIGRAM            =  0;
    public static final short               HORIZONTALORGANIGRAM        =  1;
    public static final short               TABLEHIERARCHYDIAGRAM       =  2;
    public static final short               ORGANIGRAM                  =  3;
    public static final short               VENNDIAGRAM                 = 10;
    public static final short               CYCLEDIAGRAM                = 11;
    public static final short               PYRAMIDDIAGRAM              = 12;
    public static final short               TARGETDIAGRAM               = 13;



    Controller(SmartProtocolHandler smartPH, XComponentContext xContext, XFrame xFrame){
        m_SmartPH = smartPH;
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_xController = m_xFrame.getController();
        m_NumberOfPages = getNumberOfPages();
        setGui();
        addSelectionListener();
    }

    public int getNumberOfPages(){
        Object oDocument = m_xFrame.getController().getModel();
        XDrawPagesSupplier pagesSupplier = (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, oDocument);
        XDrawPages pages = pagesSupplier.getDrawPages();
        return pages.getCount();
    }

    public SmartProtocolHandler getSmartPH(){
        return m_SmartPH;
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

    public void setGroupType(short dType){
        m_GroupType = dType;
    }

    public short getGroupType(){
        return m_GroupType;
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

    public void executeGalleryDialog(){
        if(m_Gui != null)
            m_Gui.executeGalleryDialog(true);
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
            System.err.println(ex.getLocalizedMessage());
        }
        return locale;
    }

    // adjust the DiagramId during init()
    public int getCurrentDiagramId() {
        String name = getDiagram().getShapeName(getSelectedShape());
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length() &&  charName[i] != '-')
           s +=  charName[i++];
        //return s;
        return parseInt(s);
    }

    public int getShapeID(String name){
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  charName[i] != '-')
            i++;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length())
           s +=  charName[i++];
        return parseInt(s);
    }

    public int parseInt(String s) {
        int n = -1;
        try{
            n = Integer.parseInt(s);
        }catch(NumberFormatException ex){
             System.err.println(ex.getLocalizedMessage());
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
        if(xNamed == null){
            int numOfPages = getNumberOfPages();
            // if user insert new page to the document, the control dialog will disappear
            if(numOfPages > m_NumberOfPages)
                disappearControlDialog();
            m_NumberOfPages = numOfPages;
        }else{
            m_NumberOfPages = getNumberOfPages();
            String selectedShapeName = xNamed.getName();
            // listen the diagrams
            if(selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("VennDiagram") || selectedShapeName.startsWith("PyramidDiagram") || selectedShapeName.startsWith("CycleDiagram") || selectedShapeName.startsWith("ButtDiagram")) {
                String newDiagramName = selectedShapeName.split("-", 2)[0];
                // if the previous selected item is not in the same diagram,
                // need to instantiate the new diagram
                if( m_sLastDiagramName.equals("") || !m_sLastDiagramName.equals(newDiagramName)){
                    if( selectedShapeName.startsWith("OrganizationDiagram") ){
                        setGroupType(ORGANIGROUP);
                        setDiagramType(ORGANIGRAM);
                        OrgChartTree.LASTHORLEVEL = -1;
                    }
                    if( selectedShapeName.startsWith("VennDiagram") ){
                        setGroupType(RELATIONGROUP);
                        setDiagramType(VENNDIAGRAM);
                    }
                    if( selectedShapeName.startsWith("PyramidDiagram") ){
                        setGroupType(RELATIONGROUP);
                        setDiagramType(PYRAMIDDIAGRAM);
                    }
                    if( selectedShapeName.startsWith("CycleDiagram") ){
                        setGroupType(RELATIONGROUP);
                        setDiagramType(CYCLEDIAGRAM);
                    }
                    if( selectedShapeName.startsWith("ButtDiagram") ){
                        setGroupType(RELATIONGROUP);
                        setDiagramType(TARGETDIAGRAM);
                    }
                    if( selectedShapeName.startsWith("HorizontalOrganizationDiagram") ){
                        setGroupType(ORGANIGROUP);
                        setDiagramType(HORIZONTALORGANIGRAM);
                    }
                    if( selectedShapeName.startsWith("TableHierarchyDiagram") ){
                        setGroupType(ORGANIGROUP);
                        setDiagramType(TABLEHIERARCHYDIAGRAM);
                    }
                    if( selectedShapeName.startsWith("SimpleOrganizationDiagram") ){
                        setGroupType(ORGANIGROUP);
                        setDiagramType(SIMPLEORGANIGRAM);
                    }
                    instantiateDiagram();
                    m_sLastDiagramName = newDiagramName;
                    getGui().setVisibleControlDialog(true);
                    getDiagram().initDiagram();
                }
                if((selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram")) && selectedShapeName.endsWith("RectangleShape0"))
                    if(getDiagram() != null)
                        ((OrganizationChart)getDiagram()).selectShapes();
                if(!getGui().isVisibleControlDialog())
                    getGui().setVisibleControlDialog(true);
                if(m_GroupType == RELATIONGROUP && this.isOnlySimpleItemIsSelected() && getGui().isVisibleControlDialog() && getDiagram() != null && getDiagram().isBaseColorsProps())
                    getGui().setImageColorOfControlDialog(((RelationDiagram)getDiagram()).getNextColor());
                if(getGui().isShownTextFieldOfControlDialog()){
                    if(isOnlySimpleItemIsSelected())
                        getGui().enableTextFieldOfControlDialog(true);
                    else
                        getGui().enableTextFieldOfControlDialog(false);
                }
                if(getGui().isVisibleControlDialog())
                    getGui().setFocusControlDialog();
            }else{
                disappearControlDialog();
            }
            
            if(getGui() != null)
                getGui().setTextFieldOfControlDialog();
        }
    }

    public void disappearControlDialog(){
        m_sLastDiagramName = "";
            if(getGui() != null)
                if(getGui().isVisibleControlDialog())
                    getGui().setVisibleControlDialog(false);
    }

    public void instantiateDiagram(){
  
        if(!m_IsOrganiGroupConversAction && getGui() != null)
            getGui().disposePropertiesDialog();

        if(m_DiagramType == SIMPLEORGANIGRAM)
            m_Diagram = new SimpleOrgChart(this, getGui(), m_xFrame);
        if(m_DiagramType == HORIZONTALORGANIGRAM)
            m_Diagram = new HorizontalOrgChart(this, getGui(), m_xFrame);
        if(m_DiagramType == TABLEHIERARCHYDIAGRAM)
            m_Diagram = new TableHierarchyOrgChart(this, getGui(), m_xFrame);
        if(m_DiagramType == ORGANIGRAM)
            m_Diagram = new OrgChart(this, getGui(), m_xFrame);
        if(m_DiagramType == VENNDIAGRAM)
            m_Diagram = new VennDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == CYCLEDIAGRAM)
            m_Diagram = new CycleDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == PYRAMIDDIAGRAM)
            m_Diagram = new PyramidDiagram(this, getGui(), m_xFrame);
        if(m_DiagramType == TARGETDIAGRAM)
            m_Diagram = new TargetDiagram(this, getGui(), m_xFrame);
    }

    public void convert(short diagramType){
        String newDiagramTypeName = "";
        if(diagramType == SIMPLEORGANIGRAM)
            newDiagramTypeName = "SimpleOrganizationDiagram";
        if(diagramType == HORIZONTALORGANIGRAM)
            newDiagramTypeName = "HorizontalOrganizationDiagram";
        if(diagramType == TABLEHIERARCHYDIAGRAM)
            newDiagramTypeName = "TableHierarchyDiagram";
        if(diagramType == ORGANIGRAM)
            newDiagramTypeName = "OrganizationDiagram";

        if(m_GroupType == Controller.ORGANIGROUP){
            m_IsOrganiGroupConversAction = true;
            String oldDiagramTypeName = getDiagram().getDiagramTypeName();

            //store props of old diagram
            boolean isTextFit = getDiagram().isTextFitProps();
            float fFontSize = getDiagram().getFontSizeProps();
            short style = getDiagram().getStyle();
            boolean isSelectAllShape = getDiagram().isSelectedAllShapesProps();
            short sSelectedArea = getDiagram().getSeletctedAreaProps();
            boolean isModifyColors = getDiagram().isModifyColorsProps();
            short sGradientDirection = getDiagram().getGradientDirectionProps();
            boolean isGradients = getDiagram().isGradientProps();
            boolean isPreDefinedGradients = getDiagram().isPreDefinedGradientsProps();
            short   sRounded = getDiagram().getRoundedProps();
            boolean isOutline = getDiagram().isOutlineProps();
            boolean isShadow = getDiagram().isShadowProps();
            int color = getDiagram().getColorProps();
            int startColor = getDiagram().getStartColorProps();
            int endColor = getDiagram().getEndColorProps();

            OrganizationChartTree diagramTree = ((OrganizationChart)getDiagram()).getDiagramTree();
            getDiagram().renameShapes(oldDiagramTypeName, newDiagramTypeName);
            setDiagramType(diagramType);
            instantiateDiagram();

            //set old props to the new diagram
            getDiagram().setTextFitProps(isTextFit);
            getDiagram().setFontSizeProps(fFontSize);
            getDiagram().setColorProps(color);
            getDiagram().setStartColorProps(startColor);
            getDiagram().setEndColorProps(endColor);
            getDiagram().setStyle(style);
            getDiagram().setSelectedAreaProps(sSelectedArea);
            getDiagram().setGradientProps(isGradients);
            getDiagram().setGradientDirectionProps(sGradientDirection);
            ((OrganizationChart)getDiagram()).setPropertiesValues(isSelectAllShape, isModifyColors, isPreDefinedGradients, sRounded, isOutline, isShadow);

            ((OrganizationChart)getDiagram()).initDiagramTree(diagramTree);
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, getSelectedShape());
            String selectedShapeName = xNamed.getName();
            String newDiagramName = selectedShapeName.split("-", 2)[0];
            setLastDiagramName(newDiagramName);
            m_IsOrganiGroupConversAction = false;
        }
        if(m_GroupType == Controller.RELATIONGROUP){

            //store props of old diagram
            short style = getDiagram().getStyle();
            boolean isTextFit = getDiagram().isTextFitProps();
            float fFontSize = getDiagram().getFontSizeProps();
            boolean isBaseColors = getDiagram().isBaseColorsProps();
            int color = getDiagram().getColorProps();
            int startColor = getDiagram().getStartColorProps();
            int endColor = getDiagram().getEndColorProps();
            boolean isPreDefinedGradients = getDiagram().isPreDefinedGradientsProps();

            //basic settings
            getDiagram().setPreDefinedGradientsProps(false);
            getDiagram().setBaseColorsWithGradientsProps(false);

            ArrayList<ShapeData> shapeDatas = ((RelationDiagram)getDiagram()).getShapeDatas();
            ((RelationDiagram)getDiagram()).removeAllShapesFromDrawPage();
            setDiagramType(diagramType);
            instantiateDiagram();
 
            //set old props to the new diagram
            getDiagram().setTextFitProps(isTextFit);
            getDiagram().setFontSizeProps(fFontSize);
            getDiagram().setBaseColorsProps(isBaseColors);
            getDiagram().setColorProps(color);
            getDiagram().setStartColorProps(startColor);
            getDiagram().setEndColorProps(endColor);
            if(diagramType != VENNDIAGRAM){
                getDiagram().setPreDefinedGradientsProps(isPreDefinedGradients);
                getDiagram().setStyle(style);
            }

            getGui().setColorModeOfImageOfControlDialog();

            ((RelationDiagram)getDiagram()).createDiagram(shapeDatas);
            getDiagram().initDiagram();

            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, getSelectedShape());
            String selectedShapeName = xNamed.getName();
            String newDiagramName = selectedShapeName.split("-", 2)[0];
            setLastDiagramName(newDiagramName);
            
            if(getGui() != null)
                getGui().setTextFieldOfControlDialog();
        }
    }

    public void setSelectedShape(Object obj){
        try {
           m_xSelectionSupplier.select(obj);
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public boolean isOnlySimpleItemIsSelected(){
        if(getSelectedShapes().getCount() == 1){
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, getSelectedShape() );
            String selectedShapeName = xNamed.getName();
            if((selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram")) && selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0"))
                return true;
            if(selectedShapeName.startsWith("VennDiagram") && (selectedShapeName.contains("RectangleShape") || (selectedShapeName.contains("EllipseShape") && !selectedShapeName.endsWith("EllipseShape0"))))
                return true;
            if(selectedShapeName.startsWith("CycleDiagram") && (selectedShapeName.contains("RectangleShape") || (selectedShapeName.contains("ClosedBezierShape"))))
                return true;
            if(selectedShapeName.startsWith("PyramidDiagram") && selectedShapeName.contains("PolyPolygonShape") && !selectedShapeName.endsWith("PolyPolygonShape0") )
                return true;
            if(selectedShapeName.startsWith("ButtDiagram") && ((selectedShapeName.contains("EllipseShape")  && !selectedShapeName.endsWith("EllipseShape0")) || (selectedShapeName.contains("RectangleShape")  && !selectedShapeName.endsWith("RectangleShape0"))))
                return true;
        }
        return false;
    }

    public boolean isAppropriateShapeIsSelected(){
        XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, getSelectedShape());
        if(xNamed != null){
            String selectedShapeName = xNamed.getName();
            String sDiagramType = getDiagram().getDiagramTypeName();
            if(selectedShapeName.startsWith(sDiagramType))
                return true;
        }
        return false;
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
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }

     public void setTextOfSelectedShape(String str){
        if(getSelectedShapes().getCount() == 1){
            XShape xSelectedShape = getSelectedShape();
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
            String selectedShapeName = xNamed.getName();
            if((selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram")) && selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0")){
                XText xText = (XText)UnoRuntime.queryInterface(XText.class, xSelectedShape);
                xText.setString(str);
            }
            if(getGroupType() == Controller.RELATIONGROUP){
                if((selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0")) || (selectedShapeName.contains("PolyPolygonShape") && !selectedShapeName.endsWith("PolyPolygonShape0"))){
                    XText xText = (XText)UnoRuntime.queryInterface(XText.class, xSelectedShape);
                    xText.setString(str);
                }else{
                    if(selectedShapeName.contains("EllipseShape") && !selectedShapeName.endsWith("EllipseShape0") || selectedShapeName.contains("ClosedBezierShape")){
                        RelationDiagramItem item = ((RelationDiagram)getDiagram()).getItem(xSelectedShape);
                        if(item != null){
                            XShape xTextShape = item.getTextShape();
                            if(xTextShape != null && getDiagram().isInGruopShapes(xTextShape)){
                                XText xText = (XText)UnoRuntime.queryInterface(XText.class, xTextShape);
                                xText.setString(str);
                            }
                        }
                    }
                }
            }
        }else{
            //error message
        }
    }
/*
    public void testServices(Object obj){
        XServiceInfo xSI = (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, obj);
        System.out.println("---------------------- test services ----------------------------");
        System.out.println("implementation name:");
        System.out.println(xSI.getImplementationName());
        System.out.println("supported services:");
        String[] str = xSI.getSupportedServiceNames();
        for(int i = 0; i < str.length; i++)
            System.out.println(str[i]);
        System.out.println("-----------------------------------------------------------------");
    }

    public void testProps(Object obj){
        XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, obj);
        Property[] props = xProp.getPropertySetInfo().getProperties();
        System.out.println("---------------------- test properties --------------------------");
        for (Property p : props)
            System.out.println(p.Name + " "  + p.Type.getTypeName());
        System.out.println("-----------------------------------------------------------------");
    }
*/
}
