package oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
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
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;


public class PyramidDiagram extends RelationDiagram {


    protected int               m_iGroupWidth;
    protected int               m_iGroupHeight;

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_SHADOW         = 2;
    public final static short   BC_WITH_GRADIENTS   = 3;
    public final static short   BLUE_GRADIENTS      = 4;
    public final static short   RED_GRADIENTS       = 5;
    public final static short   USER_DEFINE         = 6;

    private final static int    SHADOW_DIST         = 400;
    private final static int    SHADOW_TRANSP       = 30;
    

    public PyramidDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth   = 5;
        m_iGroupHeight  = 4;

        setSelectedAllShapesProps(true);
        setModifyColorsProps(false);
        setBaseColorsProps(true);
        m_IsBaseColorsWithGradients = false;
        m_IsBlueGradients = false;
        m_IsRedGradients = false;
        m_IsShadow = false;
        m_IsOutline = true;

        m_IsAction = false;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "PyramidDiagram";
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }

    @Override
    public void createControlShape(){
        XShape xControlTriangle = createShape("PolyPolygonShape", 0);
        m_xShapes.add(xControlTriangle);
        Point a = new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
        Point b = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth / 2 + m_iHalfDiff, m_PageProps.BorderTop);
        Point c = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth + m_iHalfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
        setControlTriangleShape(xControlTriangle, a, b, c);
        setControlShapeProps(xControlTriangle);
    }
    
    @Override
    public void setDrawArea(){
        try {
            m_PageProps.BorderTop += (SHADOW_DIST / 2 + 100);
            m_DrawAreaWidth -= ( 2 * SHADOW_DIST + 100);
            m_DrawAreaHeight -= (SHADOW_DIST / 2 + 100);

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

    public void setControlTriangleShape(XShape xTriangleShape, Point a, Point b, Point c){
        try {
                Point[] points1 = new Point[3];
                points1[0] = new Point(a.X, a.Y);
                points1[1] = new Point(c.X, c.Y);
                points1[2] = new Point(b.X, b.Y);

                Point[] points2 = new Point[2];
                points2[0] = new Point(a.X, a.Y);
                points2[1] = new Point(b.X, b.Y);

                Point[][] allPoints = new Point[2][];
                allPoints[0] = points1;
                allPoints[1] = points2;

                XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTriangleShape);
                xShapeProps.setPropertyValue("PolyPolygon", allPoints);
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
    public void createItem(int shapeID, String str) {
        int color = getColor(shapeID);
        XShape xTrapezeShape = createShape("PolyPolygonShape", shapeID);
        m_xShapes.add(xTrapezeShape);
        if(m_IsBaseColorsWithGradients || m_IsBlueGradients || m_IsRedGradients){
            if(m_IsBaseColorsWithGradients){
                setColorsOfGradients(color);
                setGradient(xTrapezeShape, GradientStyle.RECT, (short)0, (short)0, (short)50, (short)50, (short)100, (short)75);
            }
            if(m_IsBlueGradients || m_IsRedGradients){
                if(shapeID > 0)
                    setGradientsStyleColors(shapeID);
                setGradient(xTrapezeShape);
            }
        }else{
            setColorOfShape(xTrapezeShape, color);
        }
        setMoveProtectOfShape(xTrapezeShape);

        PyramidDiagramItem item = new PyramidDiagramItem(this, shapeID, xTrapezeShape);
        addItem(item);
        item.setShapesProps();
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);

        getController().setSelectedShape(xTrapezeShape);
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
                if (currShapeName.contains("PolyPolygonShape")) {
                    if(currShapeName.endsWith("PolyPolygonShape0")){
                        m_xControlShape = xCurrShape;
                    }else{
                        currShapeID = getController().getShapeID(currShapeName);
                        addItem(new PyramidDiagramItem(this, currShapeID, xCurrShape));
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapeProps(getControlShape());
    }

    @Override
    public void refreshDiagram() {
        removeSingleItems();

        XShape xControlTriangleShape = getControlShape();
        int numOfItems = getNumOfItems();
        Size controlShapeSize = xControlTriangleShape.getSize();
        Point controlShapePos = xControlTriangleShape.getPosition();

        for(RelationDiagramItem item : items)
            ((PyramidDiagramItem)item).setPosition(numOfItems, controlShapeSize, controlShapePos);
    }

    public XShape getControlShape(){
        if(m_xControlShape == null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.contains("PolyPolygonShape0"))
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
    public int getSelectedShapeID() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("PyramidDiagram") && shapeName.contains("PolyPolygonShape")){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public int getTopShapeID() {
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
        try {

            for(int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("PolyPolygonShape")){
                    shapeID = getController().getShapeID(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        }catch(IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        }catch(WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
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
                    xNextSelectedShape = getControlShape();
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
        if(shapeName.contains("PolyPolygonShape"))
            return getPreviousItem().getMainShape();
        return null;
    }

    @Override
    public XShape getNextShape() {
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("PolyPolygonShape"))
            return getNextItem().getMainShape();
        return null;
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
                        if(!shapeName.endsWith("PolyPolygonShape0") && !shapeName.equals("")){
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

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, boolean isBaseColors, boolean isBaseColorsWithGradients, boolean isBlueGradients, boolean isRedGradients, boolean isOutline, boolean isShadow){
        setSelectedAllShapesProps(isSelectAllShape);
        setModifyColorsProps(isModifyColors);
        setBaseColorsProps(isBaseColors);
        setBaseColorsWithGradientsProps(isBaseColorsWithGradients);
        setBlueGradientsProps(isBlueGradients);
        setRedGradientsProps(isRedGradients);
        setOutlineProps(isOutline);
        setShadowProps(isShadow);
        getGui().setColorModeOfImageOfControlDialog();
    }

    @Override
    public void refreshShapeProperties(){
        // need to memorize members, if user would exit into propsDialog
        boolean isSelectAllShape = isSelectedAllShapesProps();
        boolean isModifyColors = isModifyColorsProps();
        boolean isBaseColors = isBaseColorsProps();
        boolean isBaseColorsWithGradients = m_IsBaseColorsWithGradients;
        boolean isBlueGradients = m_IsBlueGradients;
        boolean isRedGradients = m_IsRedGradients;
        boolean isShadow =  m_IsShadow;
        boolean isOutline = m_IsOutline;

        m_IsAction = false;

        getGui().executePropertiesDialog();

        if(m_IsAction){
            if( m_Style == DEFAULT ){
                setPropertiesValues(true, true, true, false, false, false, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITHOUT_OUTLINE){
                setPropertiesValues(true, true, true, false, false, false, false, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITH_SHADOW){
                setPropertiesValues(true, true, true, false, false, false, true, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == BC_WITH_GRADIENTS){
                setPropertiesValues(true, true, true, true, false, false, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == BLUE_GRADIENTS)
                setPropertiesValues(true, true, false, false, true, false, true, false);
            if( m_Style == RED_GRADIENTS)
                setPropertiesValues(true, true, false, false, false, true, true, false);
            if(m_Style == USER_DEFINE){
                setModifyColorsProps(getGui().isModifyColorsPropsInDiagramPropsDialog());
                if(isModifyColorsProps()){
                    setBaseColorsWithGradientsProps(false);
                    setBlueGradientsProps(false);
                    setRedGradientsProps(false);
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

            if(isSelectedAllShapesProps())
                setAllShapeProperties();
            else
                setSelectedShapesProperties();

            setModifyColorsProps(false);
        }else{
            setSelectedAllShapesProps(isSelectAllShape);
            setModifyColorsProps(isModifyColors);
            setBaseColorsProps(isBaseColors);
            //m_IsBaseColors = isBaseColors;
            m_IsBaseColorsWithGradients = isBaseColorsWithGradients;
            m_IsBlueGradients = isBlueGradients;
            m_IsRedGradients = isRedGradients;
            m_IsShadow = isShadow;
            m_IsOutline = isOutline;
        }
        m_IsAction = false;
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);

            if(isModifyColorsProps()){
                int color = -1;
                int shapeID = getController().getShapeID(getShapeName(xShape));
                if(m_IsBaseColorsWithGradients || m_IsBlueGradients || m_IsRedGradients){
                    if(m_IsBaseColorsWithGradients){
                        if(shapeID > 0)
                            setGradientsStyleColors(shapeID);
                        setGradient(xShape, GradientStyle.RECT, (short)0, (short)0, (short)50, (short)50, (short)100, (short)75);
                    }
                    if(m_IsBlueGradients || m_IsRedGradients){
                        if(shapeID > 0)
                            setGradientsStyleColors(shapeID);
                        setGradient(xShape);
                    }
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
            }

            if(m_IsShadow){
                xProp.setPropertyValue("Shadow", new Boolean(true));
                xProp.setPropertyValue("ShadowXDistance", new Integer(SHADOW_DIST));
                xProp.setPropertyValue("ShadowYDistance", new Integer(-SHADOW_DIST/2));
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

            if(m_IsOutline)
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
            else
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
  
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

    public void setColorsOfGradients(int color){
        if (color != -1) {
            setStartColorProps(color);
            setEndColorProps(color);
        }
    }

    public void setGradientsStyleColors(int shapeID){
        --shapeID;
        if(m_IsBaseColorsWithGradients){
            int color = -1;
            color = aCOLORS[shapeID % 8];
            if (color != -1) {
                setStartColorProps(color);
                setEndColorProps(color);
            }
        }
        if(m_IsBlueGradients){
            shapeID %= 8;
            if(shapeID < 4){
                setStartColorProps(aBLUE_COLORS[shapeID]);
                setEndColorProps(aBLUE_COLORS[shapeID + 1]);
            }else{
                setStartColorProps(aBLUE_COLORS[8 - shapeID]);
                setEndColorProps(aBLUE_COLORS[8 - (shapeID + 1)]);
            }
        }
        if(m_IsRedGradients){
            shapeID %= 8;
            if(shapeID < 4){
                setStartColorProps(aRED_COLORS[shapeID]);
                setEndColorProps(aRED_COLORS[shapeID + 1]);
            }else{
                setStartColorProps(aRED_COLORS[8 - shapeID]);
                setEndColorProps(aRED_COLORS[8 - (shapeID + 1)]);
            }
        }
    }

}