package oxygenoffice.extensions.smart.diagram.organizationcharts;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.Diagram;


public abstract class OrganizationChart extends Diagram{

    // rates of measure of groupShape (e.g.: 10:6)
    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    // rates of measure of rectangles
    //(e.g.: WIDTH:HORSPACE 2:1, HEIGHT:VERSPACE 4:3)
    protected int               m_iShapeWidth;
    protected int               m_iHorSpace;
    protected int               m_iShapeHeight;
    protected int               m_iVerSpace;
    
    //horizontal offset in the drawpage if needed
    protected int               m_iHalfDiff         = 0;

    //hierarhich tpyes
    public static final short   UNDERLING           = 0;
    public static final short   ASSOCIATE           = 1;

    //item hierarhich tpye in diagram
    protected short             m_sNewItemHType     = UNDERLING;

    //styles
    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   NOT_ROUNDED         = 2;
    public final static short   WITH_SHADOW         = 3;
    public final static short   BLUE_GRADIENTS      = 4;
    public final static short   RED_GRADIENTS       = 5;
    public final static short   USER_DEFINE         = 6;

    //values of propperties
    private final static int    CORNER_RADIUS1      = 350;
    private final static int    CORNER_RADIUS2      = 700;
    public final static int     SHADOW_DIST         = 300;
    private final static int    SHADOW_TRANSP       = 30;

    protected final int[]       aRED_COLORS         = { 11674146, 14160145, 16711680, 16744192, 16776960 };
    protected final int[]       aBLUE_COLORS        = { 85, 170, 255, 7573742, 11393254 };
    
   


    public OrganizationChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_sNewItemHType = UNDERLING;
        setSelectedAllShapesProps(true);
        setModifyColorsProps(false);
        m_IsGradients = false;
        m_IsBlueGradients = false;
        m_IsRedGradients = false;
        m_sRounded = Diagram.MEDIUM_ROUNDED;
        m_IsOutline = true;
        m_IsShadow = false;
    }

    public abstract void initDiagramTree(OrganizationChartTree diagramTree);

    public void setNewItemHType(short n){
        m_sNewItemHType = n;
    }

    public XShapes getShapes(){
        return m_xShapes;
    }

    public int getShapeWidth(){
        return m_iShapeWidth;
    }

    public int getHorSpace(){
        return m_iHorSpace;
    }

    public int getShapeHeight(){
        return m_iShapeHeight;
    }

    public int getVerSpace(){
        return m_iVerSpace;
    }

    public void selectShapes(){
        getController().setSelectedShape((Object)m_xShapes);
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(4);
    }

    public abstract void createDiagram(int n);

    public void setDrawArea(){
        try {
            //allow horizontal place for shadow properties
            m_PageProps.BorderTop += (SHADOW_DIST + 100);
            m_DrawAreaWidth -= ( 2 * SHADOW_DIST + 100);
            m_DrawAreaHeight -= (SHADOW_DIST + 100);

            int orignGSWidth = m_DrawAreaWidth;
            if ((m_DrawAreaWidth / m_iGroupWidth) <= (m_DrawAreaHeight / m_iGroupHeight)) {
                m_DrawAreaHeight = m_DrawAreaWidth * m_iGroupHeight / m_iGroupWidth;
            } else {
                m_DrawAreaWidth = m_DrawAreaHeight * m_iGroupWidth / m_iGroupHeight;
            }
            // set new size of m_xGroupShape for Organigram
            m_xGroupShape.setSize(new Size(m_DrawAreaWidth, m_DrawAreaHeight));
            m_iHalfDiff = 0;
            if (orignGSWidth > m_DrawAreaWidth)
                m_iHalfDiff = (orignGSWidth - m_DrawAreaWidth) / 2;
            m_iHalfDiff += SHADOW_DIST;
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public int getTopShapeID(){
        int iTopShapeID = -1;
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")) {
                    shapeID = getController().getShapeID(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    public abstract OrganizationChartTree getDiagramTree();

    @Override
    public void removeShape(){
        
        XShapes xSelectedShapes = getController().getSelectedShapes();
        XShape xShape = null;
        try{
            if(xSelectedShapes != null){
                for(int i = 0; i < xSelectedShapes.getCount(); i++){
                    xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xSelectedShapes.getByIndex(i));
                    if(xShape != null)
                        removeShape(xShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void removeShape(XShape xSelectedShape){
        if(xSelectedShape != null){
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
            String selectedShapeName = xNamed.getName();
            if(selectedShapeName.contains("RectangleShape") && !selectedShapeName.contains("RectangleShape0")){

                if( selectedShapeName.endsWith("RectangleShape1") ){
                    String title = getGui().getDialogPropertyValue("Strings", "RoutShapeRemoveError.Title");
                    String message = getGui().getDialogPropertyValue("Strings", "RoutShapeRemoveError.Message");
                    getGui().showMessageBox(title, message);
                }else{
                    // clear everythin under the item in the tree
                    OrganizationChartTreeItem selectedItem = getDiagramTree().getTreeItem(xSelectedShape);

                    OrganizationChartTreeItem dadItem = selectedItem.getDad();
                    if(selectedItem.equals(dadItem.getFirstChild())){
                        if(selectedItem.getFirstSibling() != null){
                            dadItem.setFirstChild(selectedItem.getFirstSibling());
                        }else{
                            dadItem.setFirstChild(null);
                        }
                    }else{
                        OrganizationChartTreeItem previousSibling = getDiagramTree().getPreviousSibling(selectedItem);
                        if(previousSibling != null){
                            if(selectedItem.getFirstSibling() != null)
                                previousSibling.setFirstSibling(selectedItem.getFirstSibling());
                            else
                                previousSibling.setFirstSibling(null);
                        }
                    }

                    XShape xDadShape = selectedItem.getDad().getRectangleShape();

                    if (selectedItem.getFirstChild() != null)
                        selectedItem.getFirstChild().removeItems();

                    XShape xConnShape = getDiagramTree().getDadConnectorShape(xSelectedShape);
                    if (xConnShape != null){
                        getDiagramTree().removeFromConnectors(xConnShape);
                        m_xShapes.remove(xConnShape);
                    }
                    getDiagramTree().removeFromRectangles(xSelectedShape);
                    m_xShapes.remove(xSelectedShape);
                    setNullSelectedItem(selectedItem);
                    getController().setSelectedShape((Object)xDadShape);
                }
            }
        }
    }

    public boolean isErrorInTree(){
        if(getDiagramTree() != null){
            getDiagramTree().setLists();
            if((getDiagramTree().getRectangleListSize() - 1) != getDiagramTree().getConnectorListSize())
                return true;
            if(!getDiagramTree().isValidConnectors())
                return true;
            if(getDiagramTree().setRootItem() != 1)
                return true;
        }
        return false;
    }

    public void repairDiagram(){
        if(getDiagramTree().getRectangleListSize() == 0)
            clearEmptyDiagramAndReCreate();
        getDiagramTree().repairTree();
        initDiagram();
    }

    public void setControlShapeProps(XShape xBaseShape){
        try {
            XPropertySet xBaseProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xBaseShape);
            xBaseProps.setPropertyValue("FillStyle", FillStyle.NONE);
            xBaseProps.setPropertyValue("LineStyle", LineStyle.NONE);
            //xBaseProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
      //xBaseProps.setPropertyValue("FillTransparence", new Integer(100));
            //xBaseProps.setPropertyValue("LineColor", new Integer(0xFFFFFF));
      //xBaseProps.setPropertyValue("LineTransparence", new Integer(100));
            xBaseProps.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void refreshDiagram(){
        getDiagramTree().refresh();
    }

    public void refreshConnectorProps(){
        getDiagramTree().refreshConnectorProps();
    }

    

    public void setGradientsStylesColors(short level){
        if(m_IsBlueGradients){
            level %= 4;
            setStartColorProps(aBLUE_COLORS[level]);
            setEndColorProps(aBLUE_COLORS[level + 1]);
        }
        if(m_IsRedGradients){
            level %= 4;
            setStartColorProps(aRED_COLORS[level]);
            setEndColorProps(aRED_COLORS[level + 1]);
        }
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        int color = -1;
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            
            if(isModifyColorsProps()){
                if(m_IsGradients || m_IsBlueGradients || m_IsRedGradients){
                    if(m_IsBlueGradients || m_IsRedGradients){
                        OrganizationChartTree diagramTree = getDiagramTree();
                        if(diagramTree != null){
                            OrganizationChartTreeItem item = diagramTree.getTreeItem(xShape);
                            if(item != null){
                                if(getController().getDiagramType() == Controller.ORGANIGRAM)
                                    setGradientsStylesColors(item.getDeepOfItem());
                                else
                                    setGradientsStylesColors(item.getLevel());
                            }
                        }
                    }
                    if(getGradientDirectionProps() == Diagram.VERTICAL)
                        setGradient(xShape);
                    else
                        setGradient(xShape, (short)900);
                }else{
                    if(getGui() != null && getGui().getControlDialogWindow() != null)
                        color = getGui().getImageColorOfControlDialog();
                    if(color == -1)
                        color = m_iColor;
                    xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                    xProp.setPropertyValue("FillColor", new Integer(color));
                }
            }

            if(m_sRounded == Diagram.NULL_ROUNDED)
                xProp.setPropertyValue("CornerRadius", new Integer(0));
            if(m_sRounded == Diagram.MEDIUM_ROUNDED)
                xProp.setPropertyValue("CornerRadius", new Integer(CORNER_RADIUS1));
            if(m_sRounded == Diagram.EXTRA_ROUNDED)
                xProp.setPropertyValue("CornerRadius", new Integer(CORNER_RADIUS2));

            if(m_IsOutline){
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineWidth", new Integer(100));
            }else{
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }
    
            if(m_IsShadow){
                xProp.setPropertyValue("Shadow", new Boolean(true));
                xProp.setPropertyValue("ShadowXDistance", new Integer(SHADOW_DIST));
                xProp.setPropertyValue("ShadowYDistance", new Integer(-SHADOW_DIST));
                xProp.setPropertyValue("ShadowTransparence", new Integer(SHADOW_TRANSP));
                int shadowColor = -1;
                if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID)
                    shadowColor = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                else{
                    if(getGradientDirectionProps() == Diagram.VERTICAL)
                        shadowColor = ((Gradient)xProp.getPropertyValue("FillGradient")).StartColor;
                    if(getGradientDirectionProps() == Diagram.HORIZONTAL)
                        shadowColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                }
                if(shadowColor == -1)
                    shadowColor = 8421504;
                xProp.setPropertyValue("ShadowColor", new Integer(shadowColor));
            }else{
                xProp.setPropertyValue("Shadow", new Boolean(false));
            }

            setFontPropertiesOfShape(xShape);

        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void setFontPropertyValues(){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, getDiagramTree().getRootItem().getRectangleShape());
            TextFitToSizeType textFit = (TextFitToSizeType)xPropText.getPropertyValue("TextFitToSize");
            if(textFit.getValue() == TextFitToSizeType.NONE_value){
                setTextFitProps(false);
            } else{
                setTextFitProps(true);
            }
            float fontSizeValue = AnyConverter.toFloat(xPropText.getPropertyValue("CharHeight"));
            setFontSizeProps(fontSizeValue);
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, boolean isBlueGradients, boolean isRedGradients, short sRounded, boolean isOutline, boolean isShadow){
        setSelectedAllShapesProps(isSelectAllShape);
        setModifyColorsProps(isModifyColors);
        setBlueGradientsProps(isBlueGradients);
        setRedGradientsProps(isRedGradients);
        setRoundedProps(sRounded);
        setOutlineProps(isOutline);
        setShadowProps(isShadow);
        getGui().setColorModeOfImageOfControlDialog();
    }
  
    @Override
    public void refreshShapeProperties(){
        
        // need to memorize members, if user would exit into propsDialog
        boolean isSelectAllShape = isSelectedAllShapesProps();
        boolean isModifyColors = isModifyColorsProps();
        boolean isGradients = m_IsGradients;
        boolean isBlueGradients = m_IsBlueGradients;
        boolean isRedGradients = m_IsRedGradients;
        short sRounded = m_sRounded;
        boolean isOutline = m_IsOutline;
        boolean isShadow = m_IsShadow;
        
        m_IsAction = false;

        getGui().executePropertiesDialog();
        
        if(m_IsAction){
            
            if( m_Style == DEFAULT )
                setPropertiesValues(true, false, false, false, Diagram.MEDIUM_ROUNDED, true, false);
            if( m_Style == WITHOUT_OUTLINE)
                setPropertiesValues(true, false, false, false, Diagram.MEDIUM_ROUNDED, false, false);
            if( m_Style == NOT_ROUNDED)
                setPropertiesValues(true, false, false, false, Diagram.NULL_ROUNDED, true, false);
            if( m_Style == WITH_SHADOW)
                setPropertiesValues(true, false, false, false, Diagram.MEDIUM_ROUNDED, true, true);
            if(m_Style == BLUE_GRADIENTS || m_Style == RED_GRADIENTS){
                setGradientProps(false);
                if(getController().getDiagramType() == Controller.HORIZONTALORGANIGRAM)
                    setGradientDirectionProps(Diagram.HORIZONTAL);
                else
                    setGradientDirectionProps(Diagram.VERTICAL);
            }
            if(m_Style == BLUE_GRADIENTS)
                setPropertiesValues(true, true, true, false, Diagram.MEDIUM_ROUNDED, true, false);
            if(m_Style == RED_GRADIENTS)
                setPropertiesValues(true, true, false, true, Diagram.MEDIUM_ROUNDED, true, false);
            if(m_Style == USER_DEFINE){
                if(isModifyColorsProps()){
                    setBlueGradientsProps(false);
                    setRedGradientsProps(false);
                    setGradientProps(getGui().isGradientPropsInDiagramPropsDialog());
                    if(m_IsGradients){
                        setStartColorProps(getGui().getImageColorOfControl(getGui().m_xStartColorImageControlOfPD));
                        setEndColorProps(getGui().getImageColorOfControl(getGui().m_xEndColorImageControlOfPD));
                    }else{
                        setColorProps(getGui().getImageColorOfControl(getGui().m_xColorImageControlOfPD));
                        getGui().setImageColorOfControlDialog(m_iColor);
                    }
                }
            }
            
            if(isSelectedAllShapesProps()){
                setAllShapeProperties();
            } else {
                if(getSeletctedAreaProps() == Diagram.SELECTED_SHAPES)
                    setSelectedShapesProperties();
                if(getSeletctedAreaProps() == Diagram.SIBLING_SHAPES)
                    setSiblingShapesProperties();
                if(getSeletctedAreaProps() == Diagram.BRANCH_SHAPES)
                    setBranchShapesProperties();
                
                setAllShapeFontMeausereProperties();
            }

            setModifyColorsProps(false);
        }else{
            setSelectedAllShapesProps(isSelectAllShape);
            setModifyColorsProps(isModifyColors);
            m_IsGradients = isGradients;
            m_IsBlueGradients = isBlueGradients;
            m_IsRedGradients = isRedGradients;
            m_sRounded = sRounded;
            m_IsOutline = isOutline;
            m_IsShadow = isShadow;
        }
        m_IsAction = false;
    }
    
    public void setAllShapeProperties(){
        try {
            XShape xCurrShape = null;
            String currShapeName = "";
            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")&& !currShapeName.endsWith("RectangleShape0")) {
                    setShapeProperties(xCurrShape,"RectangleShape");
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setAllShapeFontMeausereProperties(){
        try {
            XShape xCurrShape = null;
            String currShapeName = "";
            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")&& !currShapeName.endsWith("RectangleShape0")) {
                    setFontPropertiesOfShape(xCurrShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setSelectedShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();
         
         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setShapeProperties(xCurrShape,"RectangleShape");
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setSiblingShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();

         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setSiblingShapesProperties(xCurrShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setSiblingShapesProperties(XShape xShape){
        OrganizationChartTree diagramTree = getDiagramTree();
        if(diagramTree != null){
            OrganizationChartTreeItem selectedItem = diagramTree.getTreeItem(xShape);
            if(selectedItem != null){
                OrganizationChartTreeItem dadItem = selectedItem.getDad();
                if(dadItem != null){
                    OrganizationChartTreeItem item = dadItem.getFirstChild();
                    while(item != null){
                        setShapeProperties(item.m_xRectangleShape, "RectangleShape");
                        item = item.getFirstSibling();
                    }
                }else{
                    setShapeProperties(xShape, "RectangleShape");
                }

            }
        }
    }

    public void setBranchShapesProperties(){
         XShape xCurrShape = null;
         String currShapeName = "";
         XShapes xShapes = getController().getSelectedShapes();

         try{
            if(xShapes != null){
                for(int i = 0; i < xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.endsWith("GroupShape"))
                        setAllShapeProperties();
                    else
                        if (currShapeName.contains("RectangleShape") && !currShapeName.endsWith("RectangleShape0"))
                            setBranchShapesProperties(xCurrShape);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setBranchShapesProperties(XShape xShape){
        setShapeProperties(xShape, "RectangleShape");
        OrganizationChartTree diagramTree = getDiagramTree();
        if(diagramTree != null){
            OrganizationChartTreeItem selectedItem = diagramTree.getTreeItem(xShape);
            OrganizationChartTreeItem item = selectedItem.getFirstChild();
            if(item != null)
                item.setProperties();
        }
    }

    public void setItemProperties(XShape xRectangleShape, short level){
        setMoveProtectOfShape(xRectangleShape);
        
        setFontPropertiesOfShape(xRectangleShape);

        if(m_IsGradients || m_IsBlueGradients || m_IsRedGradients){
            if(level != -1 && (m_IsBlueGradients || m_IsRedGradients))
                setGradientsStylesColors(level);
            if(getGradientDirectionProps() == Diagram.VERTICAL)
                setGradient(xRectangleShape);
            else
                setGradient(xRectangleShape, (short)900);
        } else{
            setColorOfShape(xRectangleShape, getColor());
        }
        setShapeProperties(xRectangleShape, "RectangleShape");
    }

    public int getColor(){
        int color = -1;
        if(getGui() != null && getGui().getControlDialogWindow() != null)
            color = getGui().getImageColorOfControlDialog();
        if(color < 0)
            color = m_iColor;
        return color;
    }

    public abstract void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex);

    public abstract void setConnectorShapeProps(XShape xConnShape, Integer start, Integer end);

    public void clearEmptyDiagramAndReCreate(){
        try {
             if(m_xShapes != null){
                XShape xShape = null;
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    if (xShape != null)
                        m_xShapes.remove(xShape);
                }
            }
            createDiagram(1);
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setNullSelectedItem(OrganizationChartTreeItem item){
        item.setDad(null);
        item.setFirstChild(null);
        item.setFirstSibling(null);
    }
    
}
