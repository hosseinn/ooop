package org.openoffice.extensions.diagrams.diagram.organizationcharts;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import org.openoffice.extensions.diagrams.Controller;
import org.openoffice.extensions.diagrams.Gui;
import org.openoffice.extensions.diagrams.diagram.Diagram;


public abstract class OrganizationChart extends Diagram{

    // rates of measure of groupShape (e.g.: 10:6)
    protected int               GROUPWIDTH;
    protected int               GROUPHEIGHT;

    // rates of measure of rectangles
    //(e.g.: WIDTH:HORSPACE 2:1, HEIGHT:VERSPACE 4:3)
    protected int               WIDTH;
    protected int               HORSPACE;
    protected int               HEIGHT;
    protected int               VERSPACE;

    public static final short   UNDERLING           = 0;
    public static final short   ASSOCIATE           = 1;

    //item hierarhich tpye in diagram
    protected short             m_sNewItemHType     = UNDERLING;

    protected final short       DEFAULT             = 0;
    protected final short       NOT_MONOGRAPHIC     = 1;
    protected final short       NOT_ROUNDED         = 2;
    protected final short       GRADIENTS           = 3;
    protected final short       USER_DEFINE         = 4;

    protected int               m_iHalfDiff         = 0;
    
    protected boolean           m_IsGradientAction  = false;
    public static int           _startColor         = 16711680;
    public static int           _endColor           = 8388608;

    //private FillStyle fillStyle = FillStyle.SOLID;
    //private LineStyle lineStyle = LineStyle.SOLID;
    //private int lineWidth = 100;
    //private int cornerRadius = 600;


    public OrganizationChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_sNewItemHType = UNDERLING;
    }

    public abstract void initDiagramTree(DiagramTree diagramTree);

    public void setGradientAction(boolean bool){
        m_IsGradientAction = bool;
    }

    public void setNewItemHType(short n){
        m_sNewItemHType = n;
    }

    public XShapes getShapes(){
        return m_xShapes;
    }

    public int getWIDTH(){
        return WIDTH;
    }

    public int getHORSPACE(){
        return HORSPACE;
    }

    public int getHEIGHT(){
        return HEIGHT;
    }

    public int getVERSPACE(){
        return VERSPACE;
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
            int orignGSWidth = m_DrawAreaWidth;
            if ((m_DrawAreaWidth / GROUPWIDTH) <= (m_DrawAreaHeight / GROUPHEIGHT)) {
                m_DrawAreaHeight = m_DrawAreaWidth * GROUPHEIGHT / GROUPWIDTH;
            } else {
                m_DrawAreaWidth = m_DrawAreaHeight * GROUPWIDTH / GROUPHEIGHT;
            }
            // set new size of m_xGroupShape for Organigram
            m_xGroupShape.setSize(new Size(m_DrawAreaWidth, m_DrawAreaHeight));
            m_iHalfDiff = 0;
            if (orignGSWidth > m_DrawAreaWidth)
                m_iHalfDiff = (orignGSWidth - m_DrawAreaWidth) / 2;
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.out.println(ex.getLocalizedMessage());
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
                    shapeID = getController().getNumberOfShape(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.out.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.out.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    public abstract DiagramTree getDiagramTree();

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

    public void setPropsOfBaseShape(XShape xBaseShape){
        try {
            XPropertySet xBaseProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xBaseShape);
            //xBaseProps.setPropertyValue("FillColor", new Integer(0xFFFF00));
            xBaseProps.setPropertyValue("FillTransparence", new Integer(100));
            //xBaseProps.setPropertyValue("LineColor", new Integer(0xFFFFFF));
            xBaseProps.setPropertyValue("LineTransparence", new Integer(100));
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

    public void setTextFitToSize(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
        } catch (Exception ex) {
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

    @Override
    public void setShapeProperties(XShape xShape, String type) {

        int shapeID = getController().getNumberOfShape(getShapeName(xShape));

        int color = -1;
        if(getGui() != null && getGui().getControlDialogWindow() != null)
            color = getGui().getImageColorOfControlDialog();
        //if(color < 0)
        //    color = COLOR;
        XPropertySet xProp = null;

        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
/*
            if(m_Style == GRADIENTS)
                if(getGui().isEnableControlDialogImageColor())
                    getGui().disableControlDialogImageColor();
            if(m_Style == DEFAULT)
                if(!getGui().isEnableControlDialogImageColor())
                    getGui().enableControlDialogImageColor();
 *
 */
 
            if( m_Style != GRADIENTS || !( m_Style == USER_DEFINE && m_IsGradients == true)){
                if(!getGui().isEnableControlDialogImageColor())
                    getGui().enableControlDialogImageColor();
            }
  
            /*
            if( m_Style == DEFAULT){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                if(color < 0)
                    color = aCOLORS[(shapeID - 1) % aCOLORS.length];
                int index = -1;
                for(int i = 0; i < aCOLORS.length; i++)
                    if(aCOLORS[i] == color)
                        index = (i + 1) % aCOLORS.length;
                if(index == -1)
                    index = shapeID % aCOLORS.length;
                xProp.setPropertyValue("FillColor", new Integer(color));
                if(getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(aCOLORS[index] );
            }
             */
            if( m_Style == DEFAULT){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineWidth", new Integer(100));
                xProp.setPropertyValue("CornerRadius", new Integer(600));
            }else if( m_Style == NOT_MONOGRAPHIC){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                xProp.setPropertyValue("LineWidth", new Integer(100));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                xProp.setPropertyValue("CornerRadius", new Integer(600));
            }else if( m_Style == NOT_ROUNDED){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineWidth", new Integer(100));
                xProp.setPropertyValue("CornerRadius", new Integer(0));
            } else if( m_Style == GRADIENTS){
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineWidth", new Integer(100));
                xProp.setPropertyValue("CornerRadius", new Integer(600));

                getGui().disableControlDialogImageColor();

                xProp.setPropertyValue("FillStyle", FillStyle.GRADIENT);
                Gradient aGradient = new Gradient();
                aGradient.Style = GradientStyle.LINEAR;
                aGradient.StartColor = _startColor;
                aGradient.EndColor = _endColor;
                aGradient.Angle = 450;
                aGradient.Border = 0;
                aGradient.XOffset = 0;
                aGradient.YOffset = 0;
                aGradient.StartIntensity = 100;
                aGradient.EndIntensity = 100;
                aGradient.StepCount = 10;
                xProp.setPropertyValue("FillGradient", aGradient);
            }else if( m_Style == USER_DEFINE){
                if(!m_IsGradients){
                    xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                    xProp.setPropertyValue("FillColor", new Integer(m_iColor));
                }else{
                    getGui().disableControlDialogImageColor();
                    xProp.setPropertyValue("FillStyle", FillStyle.GRADIENT);
                    Gradient aGradient = new Gradient();
                    aGradient.Style = GradientStyle.LINEAR;
                    aGradient.StartColor = m_iStartColor;
                    aGradient.EndColor = m_iEndColor;
                    aGradient.Angle = 450;
                    aGradient.Border = 0;
                    aGradient.XOffset = 0;
                    aGradient.YOffset = 0;
                    aGradient.StartIntensity = 100;
                    aGradient.EndIntensity = 100;
                    aGradient.StepCount = 10;
                    xProp.setPropertyValue("FillGradient", aGradient);

                }

                int rounded = m_sRounded * 600;
                if(m_sRounded == 2)
                    rounded = 1000;
                xProp.setPropertyValue("CornerRadius", new Integer(rounded));
                if(m_IsMonographic)
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
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

    public void setMoveProtectOfShape(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void refreshShapeProperties(){
        try {
            if( m_Style == USER_DEFINE ){
                // need to memorize members, if user exit into propsDialog
                boolean isSelectAllShape = m_IsSelectAllShape;
                boolean isGradients = m_IsGradients;
                short sRounded = m_sRounded;
                boolean isMonographic = m_IsMonographic;

                //default values
                m_IsSelectAllShape = true;
                m_IsGradients = false;
                m_sRounded = (short)0;
                m_IsMonographic = true;

                m_IsAction = false;

                getGui().showDiagramPropsDialog();
                if(m_IsAction){
                    if(m_IsSelectAllShape){
                        setAllShapeProperties();
                    }else{
                        XShape xCurrShape = null;
                        String currShapeName = "";
                        XShapes xShapes = getController().getSelectedShapes();

                        if (xShapes != null){
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
                    }
                }else{
                    m_IsSelectAllShape = isSelectAllShape;
                    m_IsGradients = isGradients;
                    m_sRounded = sRounded;
                    m_IsMonographic  = isMonographic;
                }
                m_IsAction = false;
            }else{

                if( m_Style == GRADIENTS ){
                    getGui().disableControlDialogImageColor();
                    m_IsGradientAction = false;
                    getGui().showGradientDialog();
                }

                if( m_Style != GRADIENTS || (m_Style == GRADIENTS && m_IsGradientAction)){
                    XShape xCurrShape = null;
                    String currShapeName = "";
                    XShapes xShapes = getController().getSelectedShapes();

                    if (xShapes != null){
                        for(int i = 0; i < xShapes.getCount(); i++){
                            xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                            currShapeName = getShapeName(xCurrShape);
                            if(currShapeName.endsWith("GroupShape"))
                                setAllShapeProperties();
                            else
                                if (currShapeName.contains("RectangleShape")&& !currShapeName.endsWith("RectangleShape0"))
                                    setShapeProperties(xCurrShape,"RectangleShape");
                        }
                    }
                    m_IsGradientAction = false;
                }
            }
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }

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

    public void setNullSelectedItem(TreeItem item){
        item.setDad(null);
        item.setFirstChild(null);
        item.setFirstSibling(null);
    }
    
}
