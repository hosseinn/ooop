package oxygenoffice.extensions.smart;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XNamed;
import com.sun.star.document.XExporter;
import com.sun.star.document.XFilter;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XStorable;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionChangeListener;
import com.sun.star.view.XSelectionSupplier;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.diagram.DataOfDiagram;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.ShapeData;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.horizontalorgchart.HorizontalOrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart.OrgChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.simpleorgchart.SimpleOrgChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.tablehierarchyorgchart.TableHierarchyOrgChart;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;
import oxygenoffice.extensions.smart.diagram.processes.bendingprocess.BendingProcess;
import oxygenoffice.extensions.smart.diagram.processes.continuousblockprocess.ContinuousBlockProcess;
import oxygenoffice.extensions.smart.diagram.processes.staggeredprocess.StaggeredProcess;
import oxygenoffice.extensions.smart.diagram.processes.upwardarrowprocess.UpwardArrowProcess;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram.CycleDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram.PyramidDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.targetdiagram.TargetDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram.VennDiagram;
import oxygenoffice.extensions.smart.gui.Gui;



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

    public static final short               ORGANIGROUP                 = 0;
    public static final short               RELATIONGROUP               = 1;
    public static final short               PROCESSGROUP                = 2;
    public static final short               LISTGROUP                   = 3;
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
    public static final short               CONTINUOUSBLOCKPROCESS      = 20;
    public static final short               STAGGEREDPROCESS            = 21;
    public static final short               BENDINGPROCESS              = 22;
    public static final short               UPWARDARROWPROCESS          = 23;

    Controller(SmartProtocolHandler smartPH, XComponentContext xContext, XFrame xFrame){
        m_SmartPH = smartPH;
        m_xContext  = xContext;
        m_xFrame    = xFrame;
        m_xController = m_xFrame.getController();
        setGui();
        addSelectionListener();
    }

    public boolean isSmARTDiagramShape(String shapeName){
        return  shapeName.startsWith("SimpleOrganizationDiagram") ||
                shapeName.startsWith("TableHierarchyDiagram") ||
                shapeName.startsWith("HorizontalOrganizationDiagram") ||
                shapeName.startsWith("OrganizationDiagram") ||
                shapeName.startsWith("VennDiagram") ||
                shapeName.startsWith("PyramidDiagram") ||
                shapeName.startsWith("CycleDiagram") ||
                shapeName.startsWith("TargetDiagram") ||
                shapeName.startsWith("ContinuousBlockProcess") ||
                shapeName.startsWith("StaggeredProcess") ||
                shapeName.startsWith("BendingProcess") ||
                shapeName.startsWith("UpwardArrowProcess");
    }
    
    public boolean isSmARTOrganigramShape(String shapeName){
        return  shapeName.startsWith("SimpleOrganizationDiagram") ||
                shapeName.startsWith("TableHierarchyDiagram") ||
                shapeName.startsWith("HorizontalOrganizationDiagram") ||
                shapeName.startsWith("OrganizationDiagram");
    }
    
    public void setNewSize(){
        getDiagram().increaseSizeProp();
        int width = 0;
        int height = 0;
        int xPos = 0;
        int yPos = 0;
        if(getDiagram().getSizeProp() == Diagram.UD_SIZE){
            width = getDiagram().getUDWidthProp();
            height = getDiagram().getUDHeightProp();
            xPos = getDiagram().getUDXPosProp();
            yPos = getDiagram().getUDYPosProp();
        }
        if(getDiagram().getSizeProp() == Diagram.FULL_SIZE){
            Size s = getDiagram().getGroupShapeSize();
            getDiagram().setUDWidthProp(s.Width);
            getDiagram().setUDHeightProp(s.Height);
            Point p = getDiagram().getGroupShapePos();
            getDiagram().setUDXPosProp(p.X);
            getDiagram().setUDYPosProp(p.Y);
            width = getDiagram().m_PageProps.Width - getDiagram().m_PageProps.BorderLeft - getDiagram().m_PageProps.BorderRight;
            height = getDiagram().m_PageProps.Height - getDiagram().m_PageProps.BorderTop - getDiagram().m_PageProps.BorderBottom;
            xPos = getDiagram().m_PageProps.BorderLeft;
            yPos = getDiagram().m_PageProps.BorderTop;
        }
        getDiagram().setGroupShapeSizeAndPos(width, height, xPos, yPos);
        getDiagram().refreshDiagram();
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

    public void setNullDiagram(){
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
//System.out.println("addSelectionChangeListener");
    }

    public void removeSelectionListener(){
        if(m_xSelectionSupplier != null)
            m_xSelectionSupplier.removeSelectionChangeListener(this);
//System.out.println("removeSelectionChangeListener");
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

    public short executeGalleryDialog(){
        if(m_Gui != null)
            return m_Gui.executeGalleryDialog();
        return 0;
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
            disappearControlDialog();
        }else{
            String selectedShapeName = xNamed.getName();
            // listen onclick on the diagrams
            if(isSmARTDiagramShape(selectedShapeName)){
                String newDiagramName = selectedShapeName.split("-", 2)[0];
                // if the previous selected item is not in the same diagram,
                // need to instantiate the new diagram
                if( m_sLastDiagramName.equals("") || !m_sLastDiagramName.equals(newDiagramName)){
//                    Diagram._setDefaultBaseColors();
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
                    if( selectedShapeName.startsWith("TargetDiagram") ){
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
                    if( selectedShapeName.startsWith("ContinuousBlockProcess") ){
                        setGroupType(PROCESSGROUP);
                        setDiagramType(CONTINUOUSBLOCKPROCESS);
                    }
                    if( selectedShapeName.startsWith("StaggeredProcess") ){
                        setGroupType(PROCESSGROUP);
                        setDiagramType(STAGGEREDPROCESS);
                    }
                    if( selectedShapeName.startsWith("BendingProcess") ){
                        setGroupType(PROCESSGROUP);
                        setDiagramType(BENDINGPROCESS);
                    }
                    if( selectedShapeName.startsWith("UpwardArrowProcess") ){
                        setGroupType(PROCESSGROUP);
                        setDiagramType(UPWARDARROWPROCESS);
                    }
                    instantiateDiagram();
                    m_sLastDiagramName = newDiagramName;
                    getDiagram().initDiagram();
                    getDiagram().initProperties();
                    getGui().setVisibleControlDialog(true);
                }
               
                if((selectedShapeName.startsWith("OrganizationDiagram") || selectedShapeName.startsWith("SimpleOrganizationDiagram") || selectedShapeName.startsWith("HorizontalOrganizationDiagram") || selectedShapeName.startsWith("TableHierarchyDiagram")) && selectedShapeName.endsWith("RectangleShape0"))
                    if(getDiagram() != null)
                        ((OrganizationChart)getDiagram()).selectShapes();
                if(!getGui().isVisibleControlDialog())
                    getGui().setVisibleControlDialog(true);
        //        if(m_GroupType == RELATIONGROUP && getDiagram().getColorModeProp() == Diagram.BASE_COLORS_MODE && this.isOnlySimpleItemIsSelected() && getGui().isVisibleControlDialog() && getDiagram() != null)
        //            getGui().setImageColorOfControlDialog(((RelationDiagram)getDiagram()).getNextColor());
//                if(getGui().isShownTextFieldOfControlDialog()){

                if(isOnlySimpleItemIsSelected())
                    getGui().enableTextFieldOfControlDialog(true);
                else
                    getGui().enableTextFieldOfControlDialog(false);
//                }
                if(getGui().isVisibleControlDialog())
                    getGui().setFocusControlDialog();
               
            }else{
                disappearControlDialog();
            }
            
            setTextFieldOfControlDialog();
        }
    }
    
    public void setTextFieldOfControlDialog(){
        if(getGui() != null)
            getGui().setTextFieldOfControlDialog();
    }
    
    public void disappearControlDialog(){
        m_sLastDiagramName = "";
            if(getGui() != null){
                if(getGui().isVisibleControlDialog())
                    getGui().setVisibleControlDialog(false);
            }
    }

    public void instantiateDiagram(){        
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
        if(m_DiagramType == CONTINUOUSBLOCKPROCESS)
            m_Diagram = new ContinuousBlockProcess(this, getGui(), m_xFrame);
        if(m_DiagramType == STAGGEREDPROCESS)
            m_Diagram = new StaggeredProcess(this, getGui(), m_xFrame);
        if(m_DiagramType == BENDINGPROCESS)
            m_Diagram = new BendingProcess(this, getGui(), m_xFrame);
        if(m_DiagramType == UPWARDARROWPROCESS)
            m_Diagram = new UpwardArrowProcess(this, getGui(), m_xFrame);
    }
    
    public void convert(){
        short exec = getGui().executeConvertDialog();
        if(exec == 1){
            short convType = getGui().getConversationType();
            if(getDiagram() != null) {
                if(getGroupType() == Controller.ORGANIGROUP){
                    short lastHorLevel = getGui().getSelectedItemPosOfConvertDialogListBox();
                    OrganizationChart orgChart = (OrganizationChart)getDiagram();
                    if(orgChart.isErrorInTree()){
                        getGui().askUserForRepair(orgChart);
                    }else{
                        if(convType == Controller.ORGANIGRAM){
                            if(getDiagramType() == Controller.ORGANIGRAM){
                                if(lastHorLevel != OrgChartTree.LASTHORLEVEL){
                                    ((OrgChartTree)((OrgChart)getDiagram()).getDiagramTree()).setLastHorLevel(lastHorLevel);
                                    getDiagram().refreshDiagram();
                                    ((OrganizationChart)getDiagram()).refreshConnectorProps();
                                    if(((OrganizationChart)getDiagram()).isHiddenRootElementProp())
                                        ((OrganizationChart)getDiagram()).getDiagramTree().getRootItem().hideElement();
                                }
                            }else{
                                OrgChartTree.LASTHORLEVEL = lastHorLevel;
                            }
                        }
                        if(convType != Controller.ORGANIGRAM || (convType == Controller.ORGANIGRAM && getDiagramType() != Controller.ORGANIGRAM)){
                            convert(convType);
                            getDiagram().refreshDiagram();
                            ((OrganizationChart)getDiagram()).refreshConnectorProps();
                            if(((OrganizationChart)getDiagram()).isHiddenRootElementProp())
                                ((OrganizationChart)getDiagram()).getDiagramTree().getRootItem().hideElement();
                        }
                    }
                }
                if(getGroupType() == Controller.RELATIONGROUP || getGroupType() == Controller.PROCESSGROUP){
                    convert(convType);
                    getDiagram().refreshDiagram();
                }
            }
        }
        getGui().disposeConvertDialog();
    }

    public void convert(short diagramType){
        if(m_GroupType == Controller.ORGANIGROUP){
            String oldDiagramTypeName = getDiagram().getDiagramTypeName();
            String newDiagramTypeName = "";
            if(diagramType == SIMPLEORGANIGRAM)
                newDiagramTypeName = "SimpleOrganizationDiagram";
            if(diagramType == HORIZONTALORGANIGRAM)
                newDiagramTypeName = "HorizontalOrganizationDiagram";
            if(diagramType == TABLEHIERARCHYDIAGRAM)
                newDiagramTypeName = "TableHierarchyDiagram";
            if(diagramType == ORGANIGRAM)
                newDiagramTypeName = "OrganizationDiagram";

            //store props of old diagram
            int iColor = getDiagram().getColorProp();
            int iStartColor = getDiagram().getStartColorProp();
            int iEndColor = getDiagram().getEndColorProp();
            short sGradientDirection = getDiagram().getGradientDirectionProp();
            boolean isOutline = getDiagram().isOutlineProp();
            int iShapesLineWidth = getDiagram().getShapesLineWidhtProp();
            int iLineColor = getDiagram().getLineColorProp();
            short sRounded = getDiagram().getRoundedProp();
            boolean isTextFit = getDiagram().isTextFitProp();
            float fFontSize = getDiagram().getFontSizeProp();
            boolean isShadow = getDiagram().isShadowProp();
            int iConnLineWidth = getDiagram().getConnectorsLineWidhtProp();
            boolean isStartArrow = getDiagram().isConnectorStartArrowProp();
            boolean isEndArrow = getDiagram().isConnectorEndArrowProp();
            int connColor = getDiagram().getConnectorColorProp();
            boolean isRootHidden = ((OrganizationChart)getDiagram()).isHiddenRootElementProp();

            OrganizationChartTree diagramTree = ((OrganizationChart)getDiagram()).getDiagramTree();
            getDiagram().renameShapes(oldDiagramTypeName, newDiagramTypeName);
            setDiagramType(diagramType);
            instantiateDiagram();
            ((OrganizationChart)getDiagram()).initDiagramTree(diagramTree);
            
            //set old props to the new diagram
            getDiagram().setDefaultProps();
            getDiagram().initColorModeAndStyle();
            getDiagram().setColorProp(iColor);
            getDiagram().setStartColorProp(iStartColor);
            getDiagram().setEndColorProp(iEndColor);
            getDiagram().setGradientDirectionProp(sGradientDirection);
            getDiagram().setOutlineProp(isOutline);
            getDiagram().setShapesLineWidthProp(iShapesLineWidth);
            getDiagram().setLineColorProp(iLineColor);
            getDiagram().setRoundedProp(sRounded);  
            getDiagram().setTextFitProp(isTextFit);
            getDiagram().setFontSizeProp(fFontSize);
            getDiagram().setShadowProp(isShadow);
            if(getDiagramType() == Controller.ORGANIGRAM){
                if(!oldDiagramTypeName.equals(newDiagramTypeName))
                    getDiagram().setConnectorTypeProp(Diagram.CONN_STANDARD);
            }else{
                if(getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    getDiagram().setConnectorTypeProp(Diagram.CONN_LINE);
                else
                    getDiagram().setConnectorTypeProp(Diagram.CONN_STANDARD);
            }
            getDiagram().setConnectorColorProp(connColor);
            if(getDiagramType() == Controller.TABLEHIERARCHYDIAGRAM)
                    getDiagram().setShownConnectorsProp(false);
            getDiagram().setConnectorsLineWidthProp(iConnLineWidth);
            getDiagram().setConnectorStartArrowProp(isStartArrow);
            getDiagram().setConnectorEndArrowProp(isEndArrow);
            ((OrganizationChart)getDiagram()).setHiddenRootElementProp(isRootHidden);
            
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, getSelectedShape());
            String selectedShapeName = xNamed.getName();
            String newDiagramName = selectedShapeName.split("-", 2)[0];
            setLastDiagramName(newDiagramName);
        }
    
        if(diagramType == VENNDIAGRAM || diagramType == CYCLEDIAGRAM || diagramType == PYRAMIDDIAGRAM || diagramType == TARGETDIAGRAM){
            ArrayList<ShapeData> shapeDatas = null;
            if(getGroupType() == RELATIONGROUP){
                shapeDatas = ((RelationDiagram)getDiagram()).getShapeDatas();
                ((RelationDiagram)getDiagram()).removeAllShapesFromDrawPage();
            }
            if(getGroupType() == PROCESSGROUP){
                shapeDatas = ((ProcessDiagram)getDiagram()).getShapeDatas();
                ((ProcessDiagram)getDiagram()).removeAllShapesFromDrawPage();
            }
            
            //store props of old diagram
            short sColorMode = getDiagram().getColorModeProp();
            if(sColorMode >= Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE && sColorMode <= Diagram.LAST_COLORTHEMEGRADIENT_MODE_VALUE){
                sColorMode = (short)(sColorMode - Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE + Diagram.FIRST_COLORTHEME_MODE_VALUE);
                for(ShapeData data : shapeDatas)
                        data.convertToColor();    
            }
            if(diagramType == Controller.VENNDIAGRAM)
                if(sColorMode >= Diagram.FIRST_COLORSCHEME_MODE_VALUE && sColorMode <= Diagram.LAST_COLORSCHEME_MODE_VALUE)
                    sColorMode = Diagram.SIMPLE_COLOR_MODE;
            if(sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE && diagramType != PYRAMIDDIAGRAM)
                sColorMode = Diagram.BASE_COLORS_MODE;
            int color = getDiagram().getColorProp();
            int startColor = getDiagram().getStartColorProp();
            int endColor = getDiagram().getEndColorProp();
            int[] colors = new int[8];
            if(sColorMode == Diagram.BASE_COLORS_MODE || sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE){
                for(int i = 0; i < Diagram._aBaseColors.length; i++)
                    colors[i] = Diagram._aBaseColors[i];
            }
            boolean isTextFit = getDiagram().isTextFitProp();
            float fFontSize = getDiagram().getFontSizeProp();
            if(diagramType == Controller.PYRAMIDDIAGRAM){
                isTextFit = false;
                fFontSize = (float)32.0;
            }
            if(diagramType == Controller.VENNDIAGRAM)
                isTextFit = true;
            
            boolean isOutline = getDiagram().isOutlineProp();
            int iShapesLineWidth = getDiagram().getShapesLineWidhtProp();
            int iLineColor = getDiagram().getLineColorProp();
            //boolean bFrame = getDiagram().isFrameProp();
            //boolean bRoundedFrame = getDiagram().isRoundedFrameProp();
            //boolean isShadow = getDiagram().isShadowProp();
          
            setGroupType(RELATIONGROUP);
            setDiagramType(diagramType);
            instantiateDiagram();

            //set old props to the new diagram   
            getDiagram().setColorModeProp(sColorMode);
            getDiagram().setColorProp(color);
            getDiagram().setStartColorProp(startColor);
            getDiagram().setEndColorProp(endColor);
            if(sColorMode == Diagram.BASE_COLORS_MODE || sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE){
                for(int i = 0; i < Diagram._aBaseColors.length; i++)
                    Diagram._aBaseColors[i] =  colors[i];
            }
            getDiagram().setTextFitProp(isTextFit);
            getDiagram().setFontSizeProp(fFontSize);
            getDiagram().setOutlineProp(isOutline);
            getDiagram().setShapesLineWidthProp(iShapesLineWidth);
            getDiagram().setLineColorProp(iLineColor);
            //getDiagram().setFrameProp(bFrame);
            //getDiagram().setRoundedFrameProp(bRoundedFrame);
            //getDiagram().setShadowProp(isShadow);
            
            ((RelationDiagram)getDiagram()).createDiagram(shapeDatas);
            //!!! a controlShape-ben nem állítja be a textFit tulajdonságot. Ellenőrizni kell!!!
            getDiagram().initDiagram();
           
            getDiagram().setColorModeProp(sColorMode);         
            getDiagram().setStyleProp(((RelationDiagram)getDiagram()).getUserDefineStyleValue());

            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, getSelectedShape());
            String selectedShapeName = xNamed.getName();
            String newDiagramName = selectedShapeName.split("-", 2)[0];
            setLastDiagramName(newDiagramName);
            
            if(getGui() != null)
                getGui().setTextFieldOfControlDialog(); 
        }
        
        if(diagramType == CONTINUOUSBLOCKPROCESS || diagramType == STAGGEREDPROCESS || diagramType == BENDINGPROCESS || diagramType == UPWARDARROWPROCESS){
            ArrayList<ShapeData> shapeDatas = null;
            if(getGroupType() == RELATIONGROUP){
                shapeDatas = ((RelationDiagram)getDiagram()).getShapeDatas();
                ((RelationDiagram)getDiagram()).removeAllShapesFromDrawPage();
            }
            if(getGroupType() == PROCESSGROUP){
                shapeDatas = ((ProcessDiagram)getDiagram()).getShapeDatas();
                ((ProcessDiagram)getDiagram()).removeAllShapesFromDrawPage();
            }
            
            //store props of old diagram
            short sColorMode = getDiagram().getColorModeProp();
            int startColor = getDiagram().getStartColorProp();
            int endColor = getDiagram().getEndColorProp();
            if(sColorMode >= Diagram.FIRST_COLORTHEME_MODE_VALUE && sColorMode <= Diagram.LAST_COLORTHEME_MODE_VALUE){
                    short s = (short)(sColorMode - Diagram.FIRST_COLORTHEME_MODE_VALUE);
                    if(s >= 0 && s < 10){
                        startColor = Diagram.aLOCOLORS[s][0];
                        endColor = Diagram.aLOCOLORS[s][1];
                    }
                    sColorMode = (short)(sColorMode - Diagram.FIRST_COLORTHEME_MODE_VALUE + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
                    for(ShapeData data : shapeDatas)
                        data.convertToGradient(startColor, endColor);
            }
            if(sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE && diagramType != UPWARDARROWPROCESS)
                sColorMode = Diagram.BASE_COLORS_MODE;
            int color = getDiagram().getColorProp();
            int[] colors = new int[8];
            if(sColorMode == Diagram.BASE_COLORS_MODE || sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE){
                for(int i = 0; i < Diagram._aBaseColors.length; i++)
                    colors[i] = Diagram._aBaseColors[i];
            }
            boolean isTextFit = getDiagram().isTextFitProp();
            float fFontSize = getDiagram().getFontSizeProp();
            if(diagramType == STAGGEREDPROCESS || diagramType == UPWARDARROWPROCESS){
                isTextFit = false;
                fFontSize = (float)32.0;
            }
            
            boolean isOutline = getDiagram().isOutlineProp();
            int iShapesLineWidth = getDiagram().getShapesLineWidhtProp();
            int iLineColor = getDiagram().getLineColorProp();
            //short sRounded = getDiagram().getRoundedProp();
            //boolean isShadow = getDiagram().isShadowProp();
         
            setGroupType(PROCESSGROUP);
            setDiagramType(diagramType);
            instantiateDiagram();

            //set old props to the new diagram   
            getDiagram().setColorModeProp(sColorMode);
            getDiagram().setColorProp(color);
            getDiagram().setStartColorProp(startColor);
            getDiagram().setEndColorProp(endColor);
            if(sColorMode == Diagram.BASE_COLORS_MODE || sColorMode == Diagram.BASE_COLORS_WITH_GRADIENT_MODE){
                for(int i = 0; i < Diagram._aBaseColors.length; i++)
                    Diagram._aBaseColors[i] =  colors[i];
            }
            getDiagram().setTextFitProp(isTextFit);
            getDiagram().setFontSizeProp(fFontSize);
            getDiagram().setOutlineProp(isOutline);
            getDiagram().setShapesLineWidthProp(iShapesLineWidth);
            getDiagram().setLineColorProp(iLineColor);
            if(diagramType == STAGGEREDPROCESS || diagramType == BENDINGPROCESS)
                getDiagram().setRoundedProp(Diagram.NULL_ROUNDED);
            //getDiagram().setShadowProp(isShadow);
            //if(diagramType == UPWARDARROWPROCESS)
            //    getDiagram().setShadowProp(false);
            
            if(diagramType == CONTINUOUSBLOCKPROCESS || diagramType == UPWARDARROWPROCESS){
                int arrowColor = getDiagram().getArrowColor();
                if(diagramType == CONTINUOUSBLOCKPROCESS)
                    ((ContinuousBlockProcess)getDiagram()).setArrowShapeColorProp(arrowColor);
                if(diagramType == UPWARDARROWPROCESS)
                    ((UpwardArrowProcess)getDiagram()).setArrowShapeColorProp(arrowColor);
            }

            ((ProcessDiagram)getDiagram()).createDiagram(shapeDatas);
            
            //!!! a controlShape-ben nem állítja be a textFit tulajdonságot. Ellenőrizni kell!!!
            getDiagram().initDiagram();
           
            getDiagram().setColorModeProp(sColorMode);         
            getDiagram().setStyleProp(((ProcessDiagram)getDiagram()).getUserDefineStyleValue());

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
            if(selectedShapeName.startsWith("TargetDiagram") && ((selectedShapeName.contains("EllipseShape")  && !selectedShapeName.endsWith("EllipseShape0")) || (selectedShapeName.contains("RectangleShape")  && !selectedShapeName.endsWith("RectangleShape0"))))
                return true;
            if(selectedShapeName.startsWith("ContinuousBlockProcess") && (selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0")))
                return true;
            if(selectedShapeName.startsWith("StaggeredProcess") && (((selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0"))) || selectedShapeName.contains("PolyPolygonShape")))
                return true;
            if(selectedShapeName.startsWith("BendingProcess") && (((selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0"))) || selectedShapeName.contains("PolyPolygonShape")))
                return true;
            if(selectedShapeName.startsWith("UpwardArrowProcess") && (((selectedShapeName.contains("RectangleShape") && !selectedShapeName.endsWith("RectangleShape0"))) || selectedShapeName.contains("EllipseShape")))
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
                RelationDiagramItem item = ((RelationDiagram)getDiagram()).getItem(xSelectedShape);
                if(item != null){
                    XShape xTextShape = item.getTextShape();
                    if(xTextShape != null && getDiagram().isInGruopShapes(xTextShape)){
                        XText xText = (XText)UnoRuntime.queryInterface(XText.class, xTextShape);
                        xText.setString(str);
                    }
                }
            }
            if(getGroupType() == Controller.PROCESSGROUP){
                ProcessDiagramItem item = ((ProcessDiagram)getDiagram()).getItem(xSelectedShape);
                if(item != null){
                    XShape xTextShape = item.getTextShape();
                    if(xTextShape != null && getDiagram().isInGruopShapes(xTextShape)){
                        XText xText = (XText)UnoRuntime.queryInterface(XText.class, xTextShape);
                        xText.setString(str);
                    }
                }
            }
        }else{
            //error message
        }
    }
    
    public void createDiagram(){
        removeSelectionListener();
        instantiateDiagram();
        if(getDiagram() != null){
            getDiagram().createDiagram();
            //have to init object tree in organigrams
            if(getGroupType() == ORGANIGROUP)
                getDiagram().initDiagram();
            getGui().setVisibleControlDialog(true);
            if(getGui().isShownTips())
                getGui().showTipsMessageBox();
        }
        addSelectionListener();
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
    
    public void createDiagram(DataOfDiagram datas){
        //datas.print();
        //System.out.println("Root element: " + datas.isOneFirstLevelData());
        instantiateDiagram();
        if(getDiagram() != null){
            getDiagram().createDiagram(datas);
            //have to init object tree in organigrams
            if(getGroupType() == ORGANIGROUP)
                getDiagram().initDiagram();
            getGui().setVisibleControlDialog(true);
            if(getGui().isShownTips())
                getGui().showTipsMessageBox();
        }
    }

    public void createDiagramFromList(){
        removeSelectionListener();
        if(getSelectedShape() != null && getSelectedShapes().getCount() == 1 && isShapeService(getSelectedShape())){
            try {
                XText xText = (XText) UnoRuntime.queryInterface(XText.class, getSelectedShape());
                if(xText.getString().equals("") || xText.getString().trim().equals("")){
                    String title = getGui().getDialogPropertyValue("Strings2", "Strings2.CouldnotCreateDiagram.Title.Label");
                    String message = getGui().getDialogPropertyValue("Strings2", "Strings2.CouldnotCreateDiagram.Message.Label");
                    getGui().showMessageBox(title, message);
                }else{
                    DataOfDiagram datas = new DataOfDiagram();
                    XEnumeration xParagraphEnumeration = null;
                    XTextContent xTextElement = null;
                    XEnumerationAccess xEnumerationAccess = (XEnumerationAccess) UnoRuntime.queryInterface(XEnumerationAccess.class, xText);
                    xParagraphEnumeration = xEnumerationAccess.createEnumeration();
                    short m = 0;
                    while (xParagraphEnumeration.hasMoreElements()) {
                        Object o = xParagraphEnumeration.nextElement();
                        XServiceInfo xInfo = (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, o);
                        if (xInfo.supportsService("com.sun.star.text.Paragraph")) {
                            xTextElement = (XTextContent) UnoRuntime.queryInterface(XTextContent.class, o);
                            XTextRange xTextRange = xTextElement.getAnchor();
                            String[] str = xTextRange.getString().split("\n");
                            String elementContent = "";
                            if(str.length > m)
                                elementContent = str[m];
                            XPropertySet xSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xInfo);
                            short level = 0;
                            try{
                                level = AnyConverter.toShort(xSet.getPropertyValue("NumberingLevel"));
                            }catch(Exception ex){ }
                            datas.add(level, elementContent);
                            m++;
                        }
                    }

                    if(!datas.isEmpty()){
//                        datas.print();
                        short exec = executeGalleryDialog();
                        if(exec == 1){
                            Object oDocument = m_xFrame.getController().getModel();
                            XDrawPagesSupplier pagesSupplier = (XDrawPagesSupplier)UnoRuntime.queryInterface(XDrawPagesSupplier.class, oDocument);
                            XDrawPages pages = pagesSupplier.getDrawPages();
                            XDrawPage xCurrentPage = getCurrentPage();
                            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xCurrentPage);
                            short numOfPage = AnyConverter.toShort(xProp.getPropertyValue("Number"));
                            XDrawPage xNewPage = pages.insertNewByIndex(numOfPage-1);
                            XDrawView xDrawView = (XDrawView)UnoRuntime.queryInterface(XDrawView.class, m_xController);
                            xDrawView.setCurrentPage(xNewPage);
                            createDiagram(datas);
                        }
                    }else{
                        String title = getGui().getDialogPropertyValue("Strings2", "Strings2.CouldnotCreateDiagram.Title.Label");
                        String message = getGui().getDialogPropertyValue("Strings2", "Strings2.CannotCreateDiagram.Message.Label");
                        getGui().showMessageBox(title, message);
                    }
                }
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (NoSuchElementException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }else{
            String title = getGui().getDialogPropertyValue("Strings2", "Strings2.CouldnotCreateDiagram.Title.Label");
            String message = getGui().getDialogPropertyValue("Strings2", "Strings2.Help.Message.Label");
            getGui().showMessageBox(title, message);
        }
        addSelectionListener();
    }
    
    public void export(){
        String sURL = getGui().raiseSaveAsDialog();
        String ext = sURL.substring(sURL.length() - 3);
        try {
            Object GraphicExportFilter = m_xContext.getServiceManager().createInstanceWithContext("com.sun.star.drawing.GraphicExportFilter", m_xContext);
            XExporter xExporter = (XExporter)UnoRuntime.queryInterface(XExporter.class, GraphicExportFilter);
            XComponent xComp = (XComponent)UnoRuntime.queryInterface(XComponent.class, getDiagram().getGroupShape());
            xExporter.setSourceDocument(xComp);
            PropertyValue aProps[] = new PropertyValue[2];
            aProps[0] = new PropertyValue(); 
            aProps[0].Name = "MediaType";
            aProps[0].Value = "image/" + ext;
            aProps[1] = new PropertyValue();
            aProps[1].Name = "URL";
            aProps[1].Value = sURL;
            XFilter xFilter = (XFilter)UnoRuntime.queryInterface(XFilter.class, GraphicExportFilter);
            xFilter.filter(aProps);
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public String getSpace(short s){
        String  str = "";
        for(int i = 0; i < s; i++)
            str += "  ";
        return str;
    }

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
}
