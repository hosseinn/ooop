package oxygenoffice.extensions.smart.diagram.relationdiagrams.buttdiagram;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.ConnectorType;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.GradientDefinitions;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.Color;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;


public class TargetDiagram extends RelationDiagram{

    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    private XShape              m_xControlShape2 = null;

    private boolean             m_IsLeftLayout;

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_FRAME          = 2;
    public final static short   LEFT_LAYOUT         = 3;
    // 4 - 15 Gradients from GradientDefinitions.java
    public final static short   USER_DEFINE         = 16;


    public TargetDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth   = 8;
        m_iGroupHeight  = 6;

        //default values
        setSelectedAllShapesProps(true);
        m_IsLeftLayout = false;
        setModifyColorsProps(false);
        setBaseColorsProps(true);
        m_IsPreDefinedGradients = false;
        m_IsOutline = true;
        m_IsFrame = false;

        m_IsAction = false;
    }

    public boolean isLeftLayoutProperty(){
        return m_IsLeftLayout;
    }

    public void setLeftLayoutProperty(boolean bool){
        m_IsLeftLayout = bool;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "ButtDiagram";
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }

    @Override
    public void createControlShape(){
        int controlEllipseSize = m_DrawAreaHeight * 7 / 8;
        Point middlePoint = new Point(m_PageProps.BorderLeft + m_iHalfDiff + controlEllipseSize / 2, m_PageProps.BorderTop + m_DrawAreaHeight - controlEllipseSize / 2);
        XShape xControlEllipseShape = createShape("EllipseShape", 0, middlePoint.X - controlEllipseSize / 2, middlePoint.Y - controlEllipseSize / 2, controlEllipseSize, controlEllipseSize);
        m_xShapes.add(xControlEllipseShape);
        setControlShapeProps(xControlEllipseShape);

        Point controlRectPoint = new Point(m_PageProps.BorderLeft + m_iHalfDiff + controlEllipseSize + (m_DrawAreaWidth - controlEllipseSize) * 1 / 4, m_PageProps.BorderTop);
        XShape xControlRectShape = createShape("RectangleShape", 0, controlRectPoint.X, controlRectPoint.Y, (m_DrawAreaWidth - controlEllipseSize) * 3 / 4, m_DrawAreaHeight - controlEllipseSize / 2 + 250);
        m_xShapes.add(xControlRectShape);
        setControlShapeProps(xControlRectShape);
    }

    @Override
    public void setDrawArea(){
        try {
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
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void addShape(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0)
            increaseItemsIDs(selectedShapeID);
        if(selectedShapeID == -1)
            selectedShapeID = getTopShapeID();
        if(selectedShapeID == -1 ){
            super.createDiagram();
            createDiagram(1);
            getGui().setColorModeOfImageOfControlDialog();
        }else{
            createItem(selectedShapeID + 1);
            if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
                getGui().setImageColorOfControlDialog(aCOLORS[ (selectedShapeID + 1) % 8] );
        }
    }



    @Override
    public RelationDiagramItem createItem(int shapeID, String str, Color oColor){
        XShape xEllipseShape = createShape("EllipseShape", shapeID);
        m_xShapes.add(xEllipseShape);
        setMoveProtectOfShape(xEllipseShape);

        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);

        XShape xConnShape = createConnectorShape(shapeID, xEllipseShape, xRectangleShape);
        setMoveProtectOfShape(xConnShape);

        TargetDiagramItem item = new TargetDiagramItem(this, shapeID, xEllipseShape, xRectangleShape, xConnShape);
        addItem(item);

        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        getController().setSelectedShape(xEllipseShape);

        return item;
    }
    
    @Override
    public RelationDiagramItem createItem(int shapeID, String str, int color) {
        XShape xEllipseShape = createShape("EllipseShape", shapeID);
        m_xShapes.add(xEllipseShape);
        if(m_IsPreDefinedGradients){
            if(shapeID > 0)
                setGradientsStyleColors(shapeID);
            setGradient(xEllipseShape, GradientStyle.RADIAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
        }else{
            setColorOfShape(xEllipseShape, color);
        }
        setMoveProtectOfShape(xEllipseShape);

        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);

        XShape xConnShape = createConnectorShape(shapeID, xEllipseShape, xRectangleShape);
        setMoveProtectOfShape(xConnShape);

        TargetDiagramItem item = new TargetDiagramItem(this, shapeID, xEllipseShape, xRectangleShape, xConnShape);
        addItem(item);

        if(str.equals(""))
            str = " ";
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);
        item.setShapesProps();

        XShape selectedShape = getController().getSelectedShape();
        if(selectedShape != null){
            String selectedShapeName = getShapeName(selectedShape);
            if(selectedShapeName.contains("RectangleShape"))
                getController().setSelectedShape(xRectangleShape);
            else
                getController().setSelectedShape(xEllipseShape);
        }else
            getController().setSelectedShape(xEllipseShape);
        
        return item;
    }

    public XShape createConnectorShape(int ID, XShape xEllipseShape, XShape xRectShape){
        XShape xConnectorShape = null;
        try{
            xConnectorShape = createShape("ConnectorShape", ID);
            m_xShapes.add(xConnectorShape);
            XPropertySet xPropSet = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            if(xPropSet != null){
                xPropSet.setPropertyValue("StartShape", xEllipseShape);
                xPropSet.setPropertyValue("EndShape", xRectShape);
                xPropSet.setPropertyValue("EndGluePointIndex", new Integer(3));
                xPropSet.setPropertyValue("LineStyle", LineStyle.SOLID);
                xPropSet.setPropertyValue("EdgeKind", ConnectorType.LINE);
            }
        }catch(Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xConnectorShape;
    }

    @Override
    public void initDiagram(){
        super.initDiagram();
        if(!items.isEmpty())
            items.clear();
        try {
            XShape xCurrShape;
            String currShapeName = "";
            int currShapeID;

            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("RectangleShape0"))
                    m_xControlShape2 = xCurrShape;
                if (currShapeName.contains("EllipseShape")) {
                    if(currShapeName.endsWith("EllipseShape0")){
                        m_xControlShape = xCurrShape;
                    }else{
                        currShapeID = getController().getShapeID(currShapeName);
                        addItem(new TargetDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID), getConnectorOfMainShape(currShapeID)));
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapeProps(getControlEllipseShape());
        setControlShapeProps(getControlRectangleShape());
    }

    @Override
    public XShape getPairOfMainShape(int id) {
        XShape xCurrShape;
        String currShapeName = "";
        int currShapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("RectangleShape")) {
                    currShapeID = getController().getShapeID(currShapeName);
                    if(currShapeID == id)
                        return xCurrShape;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }

    public XShape getConnectorOfMainShape(int id) {
        XShape xCurrShape;
        String currShapeName = "";
        int currShapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("ConnectorShape")) {
                    currShapeID = getController().getShapeID(currShapeName);
                    if(currShapeID == id)
                        return xCurrShape;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }
    
    @Override
    public void refreshDiagram() {
        removeSingleItems();

        int numOfItems = getNumOfItems();
        Size controlEllipseSize = getControlEllipseShape().getSize();
        Point controlEllipsePos = getControlEllipseShape().getPosition();
        Point middlePoint = new Point(controlEllipsePos.X + controlEllipseSize.Width / 2, controlEllipsePos.Y + controlEllipseSize.Height / 2);
        Size controlRectSize = getControlRectangleShape().getSize();
        Point controlRectPos = getControlRectangleShape().getPosition();

        for(RelationDiagramItem item : items){
            ((TargetDiagramItem)item).setPosition(numOfItems, controlEllipseSize, controlEllipsePos, middlePoint, controlRectSize, controlRectPos);
            if(isPreDefinedGradientsProps() && m_Style != USER_DEFINE)
                setGradientColor(item.getMainShape());
        }
    }

    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setGradientsStyleColors(shapeID);
            setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
        }
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("ButtDiagram") && (shapeName.contains("EllipseShape") || shapeName.contains("RectangleShape"))){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public int getTopShapeID(){
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("EllipseShape")) {
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

    public XShape getControlEllipseShape(){
        if(m_xControlShape == null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("EllipseShape0"))
                        m_xControlShape = xCurrShape;
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape;
    }

    public XShape getControlRectangleShape(){
        if(m_xControlShape2 == null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("RectangleShape0"))
                        m_xControlShape2 = xCurrShape;
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape2;
    }

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            RelationDiagramItem selectedItem = getItem(xSelectedShape);

            if(selectedItem != null){
                XShape xNextSelectedShape = null;

                if(getTopShapeID() > 1){
                    if(selectedShapeID == getTopShapeID())
                        xNextSelectedShape = getPreviousShape();
                    else
                        xNextSelectedShape = getNextShape();
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }else{
                    xNextSelectedShape = getControlEllipseShape();
                    if(isBaseColorsProps() && getGui() != null && getGui().isVisibleControlDialog())
                        getGui().setImageColorOfControlDialog(aCOLORS[0]);
                    selectedItem.removeItem();
                }

                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
            }
        }
    }

    @Override
    public XShape getPreviousShape() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("EllipseShape"))
            return getPreviousItem().getMainShape();
        if(shapeName.contains("RectangleShape"))
            return getPreviousItem().getTextShape();
        return null;
    }

    @Override
    public XShape getNextShape() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("EllipseShape"))
            return getNextItem().getMainShape();
        if(shapeName.contains("RectangleShape"))
            return getNextItem().getTextShape();
        return null;
    }

    public void setPropertiesValues(boolean isSelectAllShape, boolean isLeftLayout, boolean isModifyColors, boolean isBaseColors, boolean isPreDefinedGradients, boolean isOutline, boolean isFrame){
        setSelectedAllShapesProps(isSelectAllShape);
        m_IsLeftLayout = isLeftLayout;
        setModifyColorsProps(isModifyColors);
        setBaseColorsProps(isBaseColors);
        setPreDefinedGradientsProps(isPreDefinedGradients);
        setOutlineProps(isOutline);
        setFrameProps(isFrame);
        getGui().setColorModeOfImageOfControlDialog();
    }

    @Override
    public void refreshShapeProperties() {
        // need to memorize members, if user would exit into propsDialog
        boolean isSelectAllShape = isSelectedAllShapesProps();
        boolean isLeftLayout = m_IsLeftLayout;
        boolean isModifyColors = isModifyColorsProps();
        boolean isBaseColors = isBaseColorsProps();
        boolean isPreDefinedGradients = m_IsPreDefinedGradients;
        boolean isOutline = m_IsOutline;
        boolean isFrame = m_IsFrame;

        m_IsAction = false;

        getGui().executePropertiesDialog();

        if(m_IsAction){
            if(m_Style == DEFAULT){
                setPropertiesValues(true, false, true, true, false, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if(m_Style == WITHOUT_OUTLINE){
                setPropertiesValues(true, false, true, true, false, false, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if(m_Style == WITH_FRAME){
                setPropertiesValues(true, false, true, true, false, true, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if(m_Style == LEFT_LAYOUT){
                setPropertiesValues(true, true, true, true, false, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if(GradientDefinitions.isPreDefinedGradient(m_Style))
                setPropertiesValues(true, false, true, false, true, true, false);
            if (m_Style == USER_DEFINE) {
                setModifyColorsProps(getGui().isModifyColorsPropsInDiagramPropsDialog());
                if(isModifyColorsProps()){
                    setPreDefinedGradientsProps(false);
                    setBaseColorsProps(getGui().isBaseColorsPropsInDiagramPropsDialog());
                    if(isBaseColorsProps()){
                        getGui().setImageColorOfControlDialog(getNextColor());
                    }else{
                        setColorProps(getGui().getImageColorOfControl(getGui().m_xColorImageControlOfPD));
                        getGui().setImageColorOfControlDialog(m_iColor);
                    }
                    getGui().setColorModeOfImageOfControlDialog();
                }
            }

            if(isSelectedAllShapesProps()){
                setAllShapeProperties();
            }else{
                setSelectedShapesProperties();
                setAllShapeFontMeausereProperties();
            }
            setModifyColorsProps(false);
            
        }else{
            setSelectedAllShapesProps(isSelectAllShape);
            m_IsLeftLayout = isLeftLayout;
            setModifyColorsProps(isModifyColors);
            setBaseColorsProps(isBaseColors);
            m_IsPreDefinedGradients = isPreDefinedGradients;
            m_IsOutline = isOutline;
            m_IsFrame = isFrame;
        }
        m_IsAction = false;
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);

            if(type.equals("EllipseShape")){
                if(isModifyColorsProps()){
                    int color = -1;
                    int shapeID = getController().getShapeID(getShapeName(xShape));
                    if(m_IsPreDefinedGradients){
                        if(shapeID > 0)
                            setGradientsStyleColors(shapeID);
                        if(m_IsLeftLayout)
                            setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)30, (short)50, (short)100, (short)100);
                        else
                            setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
                    }else{
                        if(isBaseColorsProps()){
                            color = aCOLORS[(shapeID - 1) % 8];
                        }else{
                            if(getGui() != null && getGui().getControlDialogWindow() != null)
                                color = getGui().getImageColorOfControlDialog();
                            if(color == -1)
                                color = m_iColor;
                        }
                        xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                        xProp.setPropertyValue("FillColor", new Integer(color));
                    }
                }else{
                    FillStyle fillStyle = (FillStyle)xProp.getPropertyValue("FillStyle");
                    if(fillStyle == FillStyle.GRADIENT){
                        Gradient gradient = (Gradient)xProp.getPropertyValue("FillGradient");
                        if(m_IsLeftLayout)
                            setGradient(xShape, GradientStyle.RADIAL, gradient.StartColor, gradient.EndColor, (short)0, (short)0, (short)30, (short)50, (short)100, (short)100);
                        else
                            setGradient(xShape, GradientStyle.RADIAL, gradient.StartColor, gradient.EndColor, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
                    }
                }

                if(m_IsOutline)
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }

            if(type.equals("RectangleShape")){
                if(m_IsFrame){
                    setTextColorOfShape(xShape, 0x000000);
                    XPropertySet xMainShapeProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, getItem(xShape).getMainShape());
                    FillStyle mainShapeFillStyle = (FillStyle)xMainShapeProps.getPropertyValue("FillStyle");
                    if(mainShapeFillStyle == FillStyle.SOLID){
                        xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                        int color = AnyConverter.toInt(xMainShapeProps.getPropertyValue("FillColor"));
                        xProp.setPropertyValue("FillColor", new Integer(color));
                    }
                    if(mainShapeFillStyle == FillStyle.GRADIENT){
                        Gradient gradient = (Gradient)xMainShapeProps.getPropertyValue("FillGradient");
                        setGradient(xShape, GradientStyle.RECT, gradient.StartColor, gradient.EndColor, (short)0, (short)10, (short)50, (short)50, gradient.StartIntensity, gradient.EndIntensity);
                    }
                    if (m_IsOutline)
                        xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    else
                        xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }else{
                    int color = -1;
                    setInvisibleFeatures(xShape); 
                    XPropertySet xMainShapeProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, getItem(xShape).getMainShape());
                    FillStyle mainShapeFillStyle = (FillStyle)xMainShapeProps.getPropertyValue("FillStyle");
                    if(mainShapeFillStyle == FillStyle.SOLID)
                        color = AnyConverter.toInt(xMainShapeProps.getPropertyValue("FillColor"));
                    if(mainShapeFillStyle == FillStyle.GRADIENT)
                        color = ((Gradient)xMainShapeProps.getPropertyValue("FillGradient")).StartColor;
                    if(color != -1)
                        setTextColorOfShape(xShape, color);
                }
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
    
    public void setGradientsStyleColors(int shapeID){
        --shapeID;
        if(m_IsPreDefinedGradients){
            int topShapeID = getTopShapeID();
            setStartColorProps(GradientDefinitions.getPreDefinedGradient(m_Style, shapeID, topShapeID));
            setEndColorProps(GradientDefinitions.getPreDefinedGradient(m_Style, shapeID + 1, topShapeID));
        }
    }

    @Override
    public void setSelectedShapesProperties(){
        XShapes xShapes = getController().getSelectedShapes();
        XShape xShape = null;
        try {
            for(int i = 0; i < xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                if(xShape != null){
                    String shapeName = getShapeName(xShape);
                    if(shapeName.contains("GroupShape")){
                        setAllShapeProperties();
                    }else{
                        if(!shapeName.endsWith("EllipseShape0") && !shapeName.endsWith("RectangleShape0") && !shapeName.equals("")){
                            RelationDiagramItem item = getItem(xShape);
                            if(item != null)
                                item.setShapesProps();
                        }
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setInvisibleFeatures(XShape xShape){
        try {
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xShapeProps.setPropertyValue("FillStyle", FillStyle.NONE);
            xShapeProps.setPropertyValue("LineStyle", LineStyle.NONE);
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

    public void setTextColorOfShape(XShape xShape, int color){
        try {
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xShape);
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            xTextProps.setPropertyValue( "CharColor", new Integer(color));
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

}
