package oxygenoffice.extensions.smart.diagram.processes.upwardarrowprocess;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.PolygonFlags;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class UpwardArrowProcess extends ProcessDiagram {
    
    private XShape              m_xArrowShape               = null;
    private int                 m_iArrowShapeColor          = aRAINBOWCOLORS[0];

    public static int           _ArrowEndMeasure            = 2;
    public static int           _HeightMeasure              = 5;
    
    public final static short   DEFAULT                     = 0;
    public final static short   WITHOUT_OUTLINE             = 1;
    public final static short   BASECOLORS_WITH_GRADIENT    = 2;
    public final static short   GREEN_DARK_GRADIENT         = 3;
    public final static short   GREEN_BRIGHT_GRADIENT       = 4;
    public final static short   BLUE_DARK_GRADIENT          = 5;
    public final static short   BLUE_BRIGHT_GRADIENT        = 6;
    public final static short   PURPLE_DARK_GRADIENT        = 7;
    public final static short   PURPLE_BRIGHT_GRADIENT      = 8;
    public final static short   ORANGE_DARK_GRADIENT        = 9;
    public final static short   ORANGE_BRIGHT_GRADIENT      = 10;
    public final static short   YELLOW_DARK_GRADIENT        = 11;
    public final static short   YELLOW_BRIGHT_GRADIENT      = 12;
    public final static short   BLUE_SCHEME                 = 13;
    public final static short   AQUA_SCHEME                 = 14;
    public final static short   RED_SCHEME                  = 15;
    public final static short   FIRE_SCHEME                 = 16;
    public final static short   SUN_SCHEME                  = 17;
    public final static short   GREEN_SCHEME                = 18;
    public final static short   OLIVE_SCHEME                = 19;
    public final static short   PURPLE_SCHEME               = 20;
    public final static short   PINK_SCHEME                 = 21;
    public final static short   INDIAN_SCHEME               = 22;
    public final static short   MAROON_SCHEME               = 23;
    public final static short   BROWN_SCHEME                = 24;
    public final static short   USER_DEFINE                 = 25;

    public final static short   FIRST_COLORTHEMEGRADIENT_STYLE_VALUE    = 3;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE           = 13;
    
    public UpwardArrowProcess(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
        m_iGroupWidth       = 10;
        m_iGroupHeight      = 7;
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
    public boolean isColorThemeGradientStyle(short style){
        return  style == GREEN_DARK_GRADIENT || style == GREEN_BRIGHT_GRADIENT ||
                style == BLUE_DARK_GRADIENT || style == BLUE_BRIGHT_GRADIENT || 
                style == PURPLE_DARK_GRADIENT || style == PURPLE_BRIGHT_GRADIENT ||
                style == ORANGE_DARK_GRADIENT || style == ORANGE_BRIGHT_GRADIENT ||
                style == YELLOW_DARK_GRADIENT || style == YELLOW_BRIGHT_GRADIENT;
    }
    
    @Override
    public short getColorModeOfThemeGradientStyle(short style){
        return (short)(style - FIRST_COLORTHEMEGRADIENT_STYLE_VALUE + Diagram.FIRST_COLORTHEMEGRADIENT_MODE_VALUE);
    }
    
    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
    }
    
    public void setArrowShapeColorProp(int color){
        m_iArrowShapeColor = color;
    }
    
    public int getArrowShapeColorProp(){
        return m_iArrowShapeColor;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "UpwardArrowProcess";
    }
    
    @Override
    public void createDiagram(){
        super.createDiagram();
        setTextFitProp(false);
        setFontSizeProp((float)32.0);
        createDiagram(3);
    }
    
    @Override
    public void createControlShape(){
        m_xControlShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
        m_xShapes.add(m_xControlShape);
        setControlShapeProps(getControlShape());
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
                    if (currShapeName.endsWith("RectangleShape0"))
                        m_xControlShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape;
    }
    
    public void createArrowShape(boolean isNewColor){
        try{
            m_xArrowShape = createShape("ClosedBezierShape", 0);
            m_xShapes.add(m_xArrowShape);
            if(isNewColor){
                if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                    setArrowShapeColorProp(_aBaseColors[0]);
                else
                    setArrowShapeColorProp(getColorProp());
            }
            setArrowShapeProps();
            setPosOfArrowShape();
        } catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void removeArrowShape(){
        m_xShapes.remove(getArrowShape());
        m_xArrowShape = null;
    }
    
    public void setArrowShape(){
        if(m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("ClosedBezierShape0"))
                        m_xArrowShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }
    
    public XShape getArrowShape(){
        if(m_xArrowShape == null && m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("ClosedBezierShape0"))
                        m_xArrowShape = xCurrShape;
                }
            } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xArrowShape;
    }
    
    public void setArrowShapeProps(){
        try {
            if(getArrowShape() != null){
                XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xArrowShape);
                setMoveProtectOfShape(m_xArrowShape);
                int startColor = 0xffffff;
                if(isColorSchemeMode() && m_iArrowShapeColor == SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][1])
                    startColor = SchemeDefinitions.aColorSchemes[getColorModeProp() - FIRST_COLORSCHEME_MODE_VALUE][0];
                else
                    startColor = SchemeDefinitions.getGradientColor(0xffffff, m_iArrowShapeColor, 1, 4);
                setGradient(m_xArrowShape, GradientStyle.LINEAR, startColor, m_iArrowShapeColor, (short)1200, (short)0, (short)0, (short)0, (short)100, (short)100);
                xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE1));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                setZOrder(m_xArrowShape, "ClosedBezierShape");
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setPosOfArrowShape(){
        try{
            Point csPoint = getControlShape().getPosition();
            Size csSize = getControlShape().getSize();
            //imagined control ellipse properties 
            int ellipseCenterX = csPoint.X + csSize.Width;
            int ellipseWidth = csSize.Width;// - csSize.Width / 10;
            int ellipseCenterY = csPoint.Y + csSize.Height;
            int ellipseHeight = csSize.Height / (_ArrowEndMeasure + _HeightMeasure) * _HeightMeasure;
            Point ellipsePoint = new Point(ellipseCenterX, ellipseCenterY);
            Size ellipseSize = new Size(ellipseWidth, ellipseHeight);
            int sizeWidthDiff = ellipseWidth / 32;
            int sizeHeightDiff = ellipseHeight / 10;
            Size ellipseSizeBig1 = new Size(ellipseWidth + sizeWidthDiff, ellipseHeight + sizeHeightDiff);
            Size ellipseSizeBig2 = new Size(ellipseWidth + sizeWidthDiff * 2, ellipseHeight + sizeHeightDiff * 2);
            Size ellipseSizeBig3 = new Size(ellipseWidth + sizeWidthDiff * 4, ellipseHeight + sizeHeightDiff * 4);
            Size ellipseSizeSmall1 = new Size(ellipseWidth - sizeWidthDiff, ellipseHeight - sizeHeightDiff);
            Size ellipseSizeSmall2 = new Size(ellipseWidth - sizeWidthDiff * 2, ellipseHeight - sizeHeightDiff * 2);
            Size ellipseSizeSmall3 = new Size(ellipseWidth - sizeWidthDiff * 4, ellipseHeight - sizeHeightDiff * 4);
            double angle = - Math.PI;
            double angleDiff = Math.PI / 2 / 9;
       
            Point point1 = new Point((int)(ellipsePoint.X + (ellipseSize.Width) * Math.cos(angle + Math.PI / 13)),(int)(ellipsePoint.Y + (ellipseSize.Height) * Math.sin(angle + Math.PI / 13)));
            Point point2 = new Point((int)(ellipsePoint.X + ellipseSizeBig1.Width * Math.cos(angle + angleDiff * 5)),(int)(ellipsePoint.Y + ellipseSizeBig1.Height * Math.sin(angle + angleDiff * 5)));
            Point point3 = new Point((int)(ellipsePoint.X + ellipseSizeBig2.Width * Math.cos(angle + angleDiff * 8)),(int)(ellipsePoint.Y + ellipseSizeBig2.Height * Math.sin(angle + angleDiff * 8)));
            Point point4 = new Point((int)(ellipsePoint.X + ellipseSizeBig3.Width * Math.cos(angle + angleDiff * 8)),(int)(ellipsePoint.Y + ellipseSizeBig3.Height * Math.sin(angle + angleDiff * 8)));
            Point point5 = new Point((int)(ellipsePoint.X + ellipseSize.Width * Math.cos(angle + angleDiff * 9)),(int)(ellipsePoint.Y + ellipseSize.Height * Math.sin(angle + angleDiff * 9)));
            Point point6 = new Point((int)(ellipsePoint.X + ellipseSizeSmall3.Width * Math.cos(angle + angleDiff * 8)),(int)(ellipsePoint.Y + ellipseSizeSmall3.Height * Math.sin(angle + angleDiff * 8)));
            Point point7 = new Point((int)(ellipsePoint.X + ellipseSizeSmall2.Width * Math.cos(angle + angleDiff * 8)),(int)(ellipsePoint.Y + ellipseSizeSmall2.Height * Math.sin(angle + angleDiff * 8)));
            Point point8 = new Point((int)(ellipsePoint.X + ellipseSizeSmall1.Width * Math.cos(angle + angleDiff * 5)),(int)(ellipsePoint.Y + ellipseSizeSmall1.Height * Math.sin(angle + angleDiff * 5)));
           
            PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();

            Point[] pointCoords = new Point[11];
            pointCoords[0] = point1;
            pointCoords[1] = point2;
            pointCoords[2] = point2;
            pointCoords[3] = point3;
            pointCoords[4] = point4;
            pointCoords[5] = point5;
            pointCoords[6] = point6;
            pointCoords[7] = point7;
            pointCoords[8] = point8;
            pointCoords[9] = point8;
            pointCoords[10] = point1;

            Point[][] points = new Point[1][];
            points[0] = pointCoords;
            aCoords.Coordinates = points;

            PolygonFlags[] flags = new PolygonFlags[11];
           
            flags[0] = PolygonFlags.NORMAL;
            flags[1] = PolygonFlags.CONTROL;
            flags[2] = PolygonFlags.CONTROL;
            flags[3] = PolygonFlags.NORMAL;
            flags[4] = PolygonFlags.NORMAL;
            flags[5] = PolygonFlags.NORMAL;
            flags[6] = PolygonFlags.NORMAL;
            flags[7] = PolygonFlags.NORMAL;
            flags[8] = PolygonFlags.CONTROL;
            flags[9] = PolygonFlags.CONTROL;
            flags[10] = PolygonFlags.NORMAL;

            PolygonFlags[][] flagsArray = new PolygonFlags[1][];
            flagsArray[0] = flags;
            aCoords.Flags = flagsArray;

            XPropertySet xBezierShapeProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, m_xArrowShape);
            xBezierShapeProps.setPropertyValue("PolyPolygonBezier", aCoords);

        }catch(Exception ex){
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setZOrder(XShape xShape, String type){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if(type.equals("RectangleShape") || type.equals("EllipseShape"))
                xProp.setPropertyValue("ZOrder", new Integer(2));
            if(type.equals("ClosedBezierShape"))
                xProp.setPropertyValue("ZOrder", new Integer(1));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID) {
        return createItem(shapeID, " ");
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID, String str) {
        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);
        setLeftHorAdjustOfShape(xRectangleShape);
        setZOrder(xRectangleShape, "RectangleShape");
        
        XShape xEllipseShape = createShape("EllipseShape", shapeID);
        m_xShapes.add(xEllipseShape);
        setMoveProtectOfShape(xEllipseShape);
        setZOrder(xEllipseShape, "EllipseShape");
        
        UpwardArrowProcessItem item = new UpwardArrowProcessItem(this, shapeID, xRectangleShape, xEllipseShape);
        addItem(item);
        item.setText(str);
        item.setShapesProps(true);
 
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape((Object)xEllipseShape);

        return item;
    }
    
    @Override
    public ProcessDiagramItem createItem(int shapeID, String str, Color oColor){
        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);
        setLeftHorAdjustOfShape(xRectangleShape);
        setZOrder(xRectangleShape, "RectangleShape");
        
        XShape xEllipseShape = createShape("EllipseShape", shapeID);
        m_xShapes.add(xEllipseShape);
        setMoveProtectOfShape(xEllipseShape);
        setZOrder(xEllipseShape, "EllipseShape");
        
        UpwardArrowProcessItem item = new UpwardArrowProcessItem(this, shapeID, xRectangleShape, xEllipseShape);
        addItem(item);
        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        item.setLineColor();
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape((Object)xEllipseShape);
 
        return item;
    }
    
    public void setColorSettingsOfShape(XShape xShape){
        if(isColorSchemeMode() || isColorThemeGradientMode()){
            if(isColorSchemeMode()){
                int shapeID = getController().getShapeID(getShapeName(xShape));
                if(shapeID > 0)
                    setSchemesColors(shapeID);
                setLineColorProp(getDefaultLineColor());
            }
            if(isColorThemeGradientMode())
                setColorThemeGradientColors();
            setGradient(xShape, GradientStyle.ELLIPTICAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
        }else if(isBaseColorsWithGradientMode()){
            setLineColorProp(getDefaultLineColor());
            setGradient(xShape, GradientStyle.ELLIPTICAL, getColorProp(), getColorProp(), (short)0, (short)0, (short)50, (short)50, (short)100, (short)85);
        }else{
            if(isSimpleColorMode())
                setLineColorProp(getDefaultLineColor());    
            if(isBaseColorsMode())
                setLineColorProp(getDefaultLineColor());
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

            if(type.equals("BaseShape")){
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
            }
            if(type.equals("EllipseShape")){
                if(isModifyColorsProp())
                    setColorSettingsOfShape(xShape);
                if(m_IsShadow){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    int numOfItems = getNumOfItems();
                    int shadowDist = SHADOW_DIST1;
                    if(numOfItems >= 4 && numOfItems <= 6)
                        shadowDist = SHADOW_DIST1 * 5 / 6;
                    if(numOfItems >= 7 && numOfItems <= 9)
                        shadowDist =  SHADOW_DIST1 * 4 / 6;
                    if(numOfItems >= 10)
                        shadowDist = SHADOW_DIST1 / 2;
                    xProp.setPropertyValue("ShadowXDistance", new Integer(shadowDist));
                    xProp.setPropertyValue("ShadowYDistance", new Integer(-shadowDist));
                    xProp.setPropertyValue("ShadowTransparence", new Integer(SHADOW_TRANSP));
                    int shadowColor = -1;
                    if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID)
                        shadowColor = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                    else
                        shadowColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                    if(shadowColor == -1)
                        shadowColor = 8421504;
                    xProp.setPropertyValue("ShadowColor", new Integer(shadowColor));
                }else{
                    xProp.setPropertyValue("Shadow", new Boolean(false));
                }
            
                if(m_IsOutline){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                } else {
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
            }
            if(type.equals("RectangleShape")){
                xProp.setPropertyValue("FillStyle", FillStyle.NONE);
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }
            if(type.equals("PolyPolygonShape")){
                if(isModifyColorsProp()){
                    setColorOfShape(xShape, getArrowShapeColorProp());
                    setArrowShapeProps();
                }
            }
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
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
            if(selectedShapeID == 0)
                createArrowShape(false);
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
        }
    }

    @Override
    public void refreshDiagram() {
        removeSingleItems();
        int numOfItems = getNumOfItems();
        if(numOfItems > 0){
            //init
            Point controlShapePoint = getControlShape().getPosition();
            Size controlShapeSize = getControlShape().getSize();
            for(ProcessDiagramItem item : items){
                ((UpwardArrowProcessItem)item).setPosition(numOfItems, controlShapeSize, controlShapePoint);
                if(isColorSchemeMode())
                    setGradientColor(item.getSecondShape());
            }
        }
    }
    
    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setSchemesColors(shapeID);
            setGradient(xShape, GradientStyle.ELLIPTICAL, (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
        }
    }
    
    public void setSchemesColors(int shapeID){
        int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
        int topShapeID = getTopShapeID();
        setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID - 1, topShapeID));
        setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, shapeID, topShapeID));
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("UpwardArrowProcess") && (shapeName.contains("EllipseShape") || shapeName.contains("RectangleShape"))){
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
        } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            ProcessDiagramItem selectedItem = getItem(xSelectedShape);
            if(selectedItem != null){
                String selectedShapeName = getShapeName(xSelectedShape);
                XShape xNextSelectedShape = getPreviousShape();
                if(!selectedShapeName.endsWith("RectangleShape0")){
                    if(selectedShapeID == 1 && getTopShapeID() == 1){
                        xNextSelectedShape = getControlShape();
                        removeArrowShape();
                    }
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }
                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                    setColorProp(getNextColor());
            }
        }
    }
    
    @Override
    public XShape getNextShape(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("RectangleShape"))
            return getNextItem().getMainShape();
        if(shapeName.contains("EllipseShape"))
            return getNextItem().getSecondShape();
        return null;
    }
    
    @Override
    public XShape getPreviousShape(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("RectangleShape"))
            return getPreviousItem().getMainShape();
        if(shapeName.contains("EllipseShape"))
            return getPreviousItem().getSecondShape();
        return null;
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
                if(currShapeName.endsWith("ClosedBezierShape0"))
                    m_xArrowShape = xCurrShape;
                if(currShapeName.endsWith("RectangleShape0"))
                    m_xControlShape = xCurrShape;
            }
            int topShapeID = getTopShapeID();
            for(int i = 1; i <= topShapeID; i++ ){
                for(int j = 0; j < m_xShapes.getCount(); j++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("RectangleShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new UpwardArrowProcessItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID)));
                    }
                }
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapePropsWithoutTextProps(getControlShape());
    }

    @Override
    public XShape getPairOfMainShape(int id){
        XShape xCurrShape;
        String currShapeName = "";
        int currShapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("EllipseShape")) {
                    currShapeID = getController().getShapeID(currShapeName);
                    if(currShapeID == id)
                        return xCurrShape;
                }
            }
        } catch (com.sun.star.lang.IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }
    
    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfUpwardArrowProcess();
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
    
    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, boolean isOutline, int lineWidth){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
    }
    
    public void initArrowShapeColor(){
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, getArrowShape());
        try {
            FillStyle fillStyle = (FillStyle)xProps.getPropertyValue("FillStyle");
            if(fillStyle.getValue() == FillStyle.GRADIENT_value){
                Gradient gradient = (Gradient)xProps.getPropertyValue("FillGradient");
                setArrowShapeColorProp(gradient.EndColor);
            }
            if(fillStyle.getValue() == FillStyle.SOLID_value)
                setArrowShapeColorProp(AnyConverter.toInt(xProps.getPropertyValue("FillColor")));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<ProcessDiagramItem> items){
        super.initProperties(xControlShape, items);
        initArrowShapeColor();
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            if(getStyleProp() == UpwardArrowProcess.WITHOUT_OUTLINE)
                setOutlineProp(false);
            //if(getStyleProp() == UpwardArrowProcess.BASECOLORS_WITH_GRADIENT){ }
            if(getStyleProp() == UpwardArrowProcess.USER_DEFINE){ 
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));               
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
}
