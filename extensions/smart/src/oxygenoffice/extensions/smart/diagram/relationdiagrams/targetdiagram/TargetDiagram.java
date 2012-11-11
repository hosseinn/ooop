package oxygenoffice.extensions.smart.diagram.relationdiagrams.targetdiagram;

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
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class TargetDiagram extends RelationDiagram{

    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    private XShape              m_xControlShape2 = null;

    private boolean             m_IsLeftLayout;

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_FRAME          = 2;
    public final static short   LEFT_LAYOUT         = 3;
    public final static short   GREEN_DARK      = 4;
    public final static short   GREEN_BRIGHT    = 5;
    public final static short   BLUE_DARK       = 6;
    public final static short   BLUE_BRIGHT     = 7;
    public final static short   PURPLE_DARK     = 8;
    public final static short   PURPLE_BRIGHT   = 9;
    public final static short   ORANGE_DARK     = 10;
    public final static short   ORANGE_BRIGHT   = 11;
    public final static short   YELLOW_DARK     = 12;
    public final static short   YELLOW_BRIGHT   = 13;
    public final static short   BLUE_SCHEME      = 14;
    public final static short   AQUA_SCHEME      = 15;
    public final static short   RED_SCHEME       = 16;
    public final static short   FIRE_SCHEME      = 17;
    public final static short   SUN_SCHEME       = 18;
    public final static short   GREEN_SCHEME     = 19;
    public final static short   OLIVE_SCHEME     = 20;
    public final static short   PURPLE_SCHEME    = 21;
    public final static short   PINK_SCHEME      = 22;
    public final static short   INDIAN_SCHEME    = 23;
    public final static short   MAROON_SCHEME    = 24;
    public final static short   BROWN_SCHEME     = 25;
    public final static short   USER_DEFINE         = 26;

    public final static short   FIRST_COLORTHEME_STYLE_VALUE = 4;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE = 14;


    public TargetDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth   = 8;
        m_iGroupHeight  = 6;
        setLeftLayoutProperty(false);
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
        setFrameProp(false);
    }

    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
    }
    
    @Override
    public boolean isColorSchemeStyle(short style){
        return  style == BLUE_SCHEME || style == AQUA_SCHEME ||
                style == RED_SCHEME || style == FIRE_SCHEME ||
                style == SUN_SCHEME || style == GREEN_SCHEME ||
                style == OLIVE_SCHEME || style == PURPLE_SCHEME ||
                style == PINK_SCHEME || style == INDIAN_SCHEME ||
                style == MAROON_SCHEME || style == BROWN_SCHEME;
    }

    @Override
    public short getColorModeOfSchemeStyle(short style){
        return (short)(style - FIRST_COLORSCHEME_STYLE_VALUE + Diagram.FIRST_COLORSCHEME_MODE_VALUE);
    }
    
    @Override
    public boolean isColorThemeStyle(short style){
        return  style == GREEN_DARK || style == GREEN_BRIGHT ||
                style == BLUE_DARK || style == BLUE_BRIGHT || 
                style == PURPLE_DARK || style == PURPLE_BRIGHT ||
                style == ORANGE_DARK || style == ORANGE_BRIGHT ||
                style == YELLOW_DARK || style == YELLOW_BRIGHT;
    }
    
    @Override
    public short getColorModeOfThemeStyle(short style){
        return (short)(style - FIRST_COLORTHEME_STYLE_VALUE + Diagram.FIRST_COLORTHEME_MODE_VALUE);
    }

    public boolean isLeftLayoutProperty(){
        return m_IsLeftLayout;
    }

    public void setLeftLayoutProperty(boolean bool){
        m_IsLeftLayout = bool;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "TargetDiagram";
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<RelationDiagramItem> items){
        super.initProperties(xControlShape, items);
        setLeftLayoutProperty(false);
        setFrameProp(false);
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            //if(getStyleProps() == VennDiagram.DEFAULT){ }
            if(getStyleProp() == TargetDiagram.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == TargetDiagram.WITH_FRAME)
                setFrameProp(false);
            if(getStyleProp() == TargetDiagram.LEFT_LAYOUT)
                setLeftLayoutProperty(true);
            if(getStyleProp() == TargetDiagram.USER_DEFINE){
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
          
                if(items.get(1) != null){
                    Point controlShapePos = xControlShape.getPosition();
                    Point secondItemPos = items.get(1).getMainShape().getPosition();
                    if(controlShapePos.X == secondItemPos.X)
                        setLeftLayoutProperty(true);
                }
                
                xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getTextShape());
                LineStyle lineStyle = (LineStyle)xProps.getPropertyValue("LineStyle");
                if(lineStyle.getValue() == LineStyle.SOLID_value)
                    setFrameProp(true); 
            }

            setFontPropertyValues();
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, items.get(0).getTextShape());
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            setTextColorProp(AnyConverter.toInt(xTextProps.getPropertyValue("CharColor")));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }
    
    @Override
    public XShape getControlShape(){
        if(m_xControlShape == null && m_xShapes != null){
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

    @Override
    public void createControlShape(){
        int controlEllipseSize = m_DrawAreaHeight * 7 / 8;
        Point middlePoint = new Point(m_PageProps.BorderLeft + m_iHalfDiff + controlEllipseSize / 2, m_PageProps.BorderTop + m_DrawAreaHeight - controlEllipseSize / 2);
        XShape xControlEllipseShape = createShape("EllipseShape", 0, middlePoint.X - controlEllipseSize / 2, middlePoint.Y - controlEllipseSize / 2, controlEllipseSize, controlEllipseSize);
        m_xShapes.add(xControlEllipseShape);
        setControlShapeProps(getControlShape());

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
        }else{
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
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
        item.setLineColor();
        getController().setSelectedShape(xEllipseShape);

        return item;
    }
    
    @Override
    public RelationDiagramItem createItem(int shapeID, String str) {
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

        if(str.equals(""))
            str = " ";
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);
        item.setShapesProps(true);

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
            for(int i = 0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("RectangleShape0"))
                    m_xControlShape2 = xCurrShape;
                if(currShapeName.endsWith("EllipseShape0"))
                    m_xControlShape = xCurrShape;
            }
            int topShapeID = getTopShapeID();
            for(int i = 1; i <= topShapeID; i++ ){
                for(int j = 0; j < m_xShapes.getCount(); j++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("EllipseShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new TargetDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID), getConnectorOfMainShape(currShapeID)));
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapePropsWithoutTextProps(getControlEllipseShape());
        setControlShapePropsWithoutTextProps(getControlRectangleShape());
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
            if(isColorSchemeMode()){
                setGradientColor(item.getMainShape());
                setShapeProperties(item.getTextShape());
            }
        }
    }

    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setSchemesColors(shapeID);
            setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
        }
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("TargetDiagram") && (shapeName.contains("EllipseShape") || shapeName.contains("RectangleShape"))){
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
                    selectedItem.removeItem();
                }

                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode())
                    setColorProp(getNextColor());
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

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, boolean isLeftLayout, boolean isOutline, int lineWidth, boolean isFrame){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        m_IsLeftLayout = isLeftLayout;
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setFrameProp(isFrame);
    }

    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfTargetDiagram();
            if(isSelectedAllShapesProp()){
                setAllShapeProperties();
                setModifyColorsProp(false);
            } else {
                setSelectedShapesProperties();
                setModifyColorsProp(false);
                setAllShapeProperties();
            }
            
            getController().getDiagram().refreshDiagram();
        } 
        getGui().enableAndSetFocusControlDialog();
    }
    
    public void setColorSettingsOfShape(XShape xShape, String type){
        if(isColorSchemeMode()){
            int shapeID = getController().getShapeID(getShapeName(xShape));
            if(shapeID > 0)
                setSchemesColors(shapeID);
            if(type.equals("EllipseShape")){
            if(isLeftLayoutProperty())
                setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)30, (short)50, (short)100, (short)100);
            else
                setGradient(xShape, GradientStyle.RADIAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
            }
            if(type.equals("RectangleShape"))
                setGradient(xShape, (short)1800);
            setLineColorProp(getDefaultLineColor());
        }else{
            if(isSimpleColorMode())
                setLineColorProp(getDefaultLineColor());    
            if(isBaseColorsMode())
                setLineColorProp(getDefaultLineColor());
            if(isColorThemeMode())
                setColorThemeColors();
            setColorOfShape(xShape);
        }
        setLineColorOfShape(xShape);
    }
    
    @Override
    public void setShapeProperties(XShape xShape, String type) {
        setFontPropertiesOfShape(xShape);
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            
            if(type.equals("BaseShape"))
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
            
            if(type.equals("ConnectorShape") && isModifyColorsProp())
                xProp.setPropertyValue("LineColor", new Integer(getLineColorProp()));
               
            if(type.equals("EllipseShape")){
                if(isModifyColorsProp()){
                    setColorSettingsOfShape(xShape, "EllipseShape");
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
                
                if(m_IsOutline){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                } else {
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
            }
            
            if(type.equals("RectangleShape")){
                if(isFrameProp()){
                    setTextColorOfShape(xShape, getTextColorProp());
                    setColorSettingsOfShape(xShape, "RectangleShape");
                    if(m_IsOutline){
                        xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                        xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                    } else {
                        xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                    }
                }else{
                    setInvisibleFeatures(xShape);
                    if(!isTextColorChange()){
                        if(isColorSchemeMode())
                            setTextColorOfShape(xShape, getStartColorProp());
                        else
                            setTextColorOfShape(xShape, getColorProp());
                    } else {
                        setTextColorOfShape(xShape, getTextColorProp());
                    }
                }
            }

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
    
    public void setSchemesColors(int shapeID){
        //--shapeID;
        shapeID = getTopShapeID() - shapeID;
        if(isColorSchemeMode()){
            int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
            int topShapeID = getTopShapeID();
            setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID, topShapeID));
            setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID + 1, topShapeID));
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
