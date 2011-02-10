package org.openoffice.extensions.diagrams.diagram.pyramiddiagram;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
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
import org.openoffice.extensions.diagrams.Controller;
import org.openoffice.extensions.diagrams.Gui;
import org.openoffice.extensions.diagrams.diagram.Diagram;


public class PyramidDiagram extends Diagram {


    protected final short   DEFAULT                 = 0;
    protected final short   GRADIENTS               = 1;
    protected final short   BASE_COLORS             = 2;
    protected final short   BASE_COLORS_GRADIENTS   = 3;
    protected final short   USER_DEFINE             = 4;
    

    public PyramidDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
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

    private void createDiagram(int n) {

        if(m_xDrawPage != null && m_xShapes != null){
            try {

                int orignGSWidth = m_DrawAreaWidth;
                if ((m_DrawAreaWidth * 4 / 5) <= (m_DrawAreaHeight))
                    m_DrawAreaHeight = m_DrawAreaWidth * 4 / 5;
                else
                    m_DrawAreaWidth = m_DrawAreaHeight * 5 / 4;
                m_xGroupShape.setSize(new Size(m_DrawAreaWidth, m_DrawAreaHeight));
                
                int halfDiff = 0;
                if (orignGSWidth > m_DrawAreaWidth)
                    halfDiff = (orignGSWidth - m_DrawAreaWidth) / 2;
                m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderTop));
                
                XShape xControlTriangle = createShape("PolyPolygonShape", 0);
                m_xShapes.add(xControlTriangle);
                Point a = new Point(m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
                Point b = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth / 2 + halfDiff, m_PageProps.BorderTop);
                Point c = new Point(m_PageProps.BorderLeft + m_DrawAreaWidth + halfDiff, m_PageProps.BorderTop + m_DrawAreaHeight);
                setControlTriangleShape(xControlTriangle, a, b, c);

                XShape xTrapezeShape = null;
                for (int i = 1; i <= n; i++) {
                    xTrapezeShape = createShape("PolyPolygonShape", i);
                    m_xShapes.add(xTrapezeShape);
                    setColorOfShape(xTrapezeShape, COLOR);
                    setTextAndTextFitToSize(xTrapezeShape);
                }

                refreshDiagram();

                if (n == 1)
                    getController().setSelectedShape((Object) xTrapezeShape);
                else
                    getController().setSelectedShape((Object) m_xShapes);
                if (getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(COLOR);

            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    @Override
    public void initDiagram(){

        super.initDiagram();

        XShape xTopShape = getTopShape();
        int iTopShapeID = -1;
        if(xTopShape != null)
            iTopShapeID = getController().getNumberOfShape(getShapeName(xTopShape));

        int color = -1;
        //if(m_Style == BASE_COLORS)
        //    color = aCOLORS[iTopShapeID % 8];
        //else
        if(xTopShape != null)
            color = getColorOfShape(xTopShape);
//
        if( color >= 0 && getGui() != null)
            getGui().setImageColorOfControlDialog(color);
    }

    @Override
    public void addShape(){

        if(m_xDrawPage != null && m_xShapes != null){

            int iTopShapeID = getTopShapeID();
            XShape xSelectedShape = getController().getSelectedShape();
            String selectedShapeName = getShapeName(xSelectedShape);
            int newShapeID = 0;

            if(selectedShapeName.endsWith("GroupShape") || selectedShapeName.contains("PolyPolygonShape0")){
                if(iTopShapeID == -1)
                    newShapeID = 1;
                if(iTopShapeID > 0)
                    newShapeID = iTopShapeID + 1;
            }else{
                if(selectedShapeName.contains("PolyPolygonShape")){
                    newShapeID = getController().getNumberOfShape(selectedShapeName);
                    increaseShapeIDs(newShapeID);
                }
            }

            int newShapeColor = -1;
            if(getGui() != null)
                newShapeColor = getGui().getImageColorOfControlDialog();
            if(newShapeColor < 0)
                newShapeColor = COLOR;

            

            if(newShapeID > 0){
                XShape xTrapezeShape = createShape("PolyPolygonShape", newShapeID);
                m_xShapes.add(xTrapezeShape);
                //if(m_Style == GRADIENTS || m_Style == BASE_COLORS_GRADIENTS || (m_Style == USER_DEFINE && m_IsBaseColorsWithGradients))
                //    setGradientOfShape(xTrapezeShape, newShapeColor, 0xFFFFFF);
               
                //else
                //    setColorOfShape(xTrapezeShape, newShapeColor);

                setShapeProperties(xTrapezeShape);

                setTextAndTextFitToSize(xTrapezeShape);
                if(iTopShapeID == -1)
                    getController().setSelectedShape((Object)xTrapezeShape);
            }

            if(m_Style == BASE_COLORS || m_Style == BASE_COLORS_GRADIENTS || (m_Style == USER_DEFINE && m_IsBaseColors)){
                if(selectedShapeName.endsWith("PolyPolygonShape0") || selectedShapeName.endsWith("GroupShape")){
                    if(getGui()!= null)
                        getGui().setImageColorOfControlDialog(aCOLORS[newShapeID  % 8]);
                }else{
                    if(getGui()!= null){

                        getGui().setImageColorOfControlDialog(aCOLORS[newShapeID  % 8]);

                        //int selectedShapeColor = getColorOfShapeByID(getController().getNumberOfShape(getShapeName(getController().getSelectedShape())));
                        //int newColor; // new random color between selectedShape and newShapeColor
                        //do{
                        //    newColor = aCOLORS[((int)(Math.random()*8)) % 8];
                        //}while(newColor == selectedShapeColor || newColor == newShapeColor);

                        //getGui().setImageColorOfControlDialog(newColor);
                    }
                }
            }
        }
    }

    public void setGradientOfShape(XShape xShape, int startColor, int endColor){
        try{
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xShapeProps.setPropertyValue("FillStyle", FillStyle.GRADIENT);
            Gradient aGradient = new Gradient();
            aGradient.Style = GradientStyle.LINEAR;
            aGradient.StartColor = startColor;
            aGradient.EndColor = endColor;
            aGradient.Angle = 1800;
            aGradient.Border = 0;
            aGradient.XOffset = 0;
            aGradient.YOffset = 0;
            aGradient.StartIntensity = 100;
            aGradient.EndIntensity = 85;
            aGradient.StepCount = 100;
            xShapeProps.setPropertyValue("FillGradient", aGradient);
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

    public int getColorOfShapeByID(int newItemID){
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;
        try {
            for( int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")) {
                    currShapeID = getController().getNumberOfShape(currShapeName);
                    if(currShapeID == newItemID){
                        return getColorOfShape(xCurrShape);
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
             System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
             System.err.println(ex.getLocalizedMessage());
        }
        return -1;
    }

    public void increaseShapeIDs(int newShapeID){

        if(m_xShapes != null){

            XShape xCurrShape = null;
            String currShapeName = "";
            int currShapeID;

            try{

                for( int i = 0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")){
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID >= newShapeID){
                            String newShapeName = currShapeName.split("PolyPolygonShape", 2)[0] + "PolyPolygonShape" + (currShapeID + 1);
                            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xCurrShape );
                            xNamed.setName(newShapeName);
                        }
                    }
                }
            
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    @Override
    public void removeShape() {

        if(m_xDrawPage != null && m_xShapes != null){

            XShape xSelectedShape = getController().getSelectedShape();
            String selectedShapeName = getShapeName(xSelectedShape);

            if(selectedShapeName.contains("PolyPolygonShape") && !selectedShapeName.contains("PolyPolygonShape0")){

                int selectedShapeID = getController().getNumberOfShape(selectedShapeName);
                m_xShapes.remove(xSelectedShape);
                decreaseShapeIDs(selectedShapeID);

                int iTopShapeID = getTopShapeID();
                int newSelectedShapeID = selectedShapeID;
                if(newSelectedShapeID > iTopShapeID)
                    newSelectedShapeID = iTopShapeID;
                if(iTopShapeID == -1)
                    newSelectedShapeID = 0;
                selectShapeByID(newSelectedShapeID);
            }
        }
    }

    public XShape getControlTriangle(){

        if(m_xShapes != null){

            XShape xCurrShape = null;
            String currShapeName = "";

            try {

                for(int i = 0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.contains("PolyPolygonShape0"))
                         return xCurrShape;
                }

            }catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }

        return null;
    }

    public void selectShapeByID(int shapeID){

        if(m_xShapes != null){

            if(shapeID == 0)
                getController().setSelectedShape((Object)getControlTriangle());

            else{

                XShape xCurrShape = null;
                String currShapeName = "";
                int currShapeID;

                try{

                    for( int i = 0; i < m_xShapes.getCount(); i++ ){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        if(currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")){
                            currShapeID = getController().getNumberOfShape(currShapeName);
                            if(currShapeID == shapeID)
                                getController().setSelectedShape((Object)xCurrShape);
                        }
                    }

                } catch (IndexOutOfBoundsException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println(ex.getLocalizedMessage());
                }

            }
        }
    }

    public void decreaseShapeIDs(int shapeID){

        if(m_xShapes != null){

            XShape xCurrShape = null;
            String currShapeName = "";
            int currShapeID;

            try{

                for( int i = 0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")){
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID > shapeID){
                            String newShapeName = currShapeName.split("PolyPolygonShape", 2)[0] + "PolyPolygonShape" + (currShapeID - 1);
                            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xCurrShape );
                            xNamed.setName(newShapeName);
                        }
                    }
                }

            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public XShape getTopShape(){

        int iTopShapeID = -1;
        XShape xTopShape = null;

        if(m_xShapes != null){

            XShape xCurrShape = null;
            String currShapeName = "";
            int currShapeID;

            try {

                for(int i = 0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")){
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID > iTopShapeID){
                            iTopShapeID = currShapeID;
                            xTopShape = xCurrShape;
                        }
                    }
                }
            }catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return xTopShape;
    }

    public int getTopShapeID(){

        int iTopShapeID = -1;
        XShape xTopShape = getTopShape();
        if(xTopShape != null)
            iTopShapeID = getController().getNumberOfShape(getShapeName(xTopShape));
        
        return iTopShapeID;
    }

    @Override
    public void refreshDiagram(){

        XShape xControlTriangleShape = null;
        int iTopShapeID = -1;
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;

        try {

            for(int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("PolyPolygonShape")){
                    if(currShapeName.contains("PolyPolygonShape0")){
                        xControlTriangleShape = xCurrShape;
                    }else{
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID > iTopShapeID)
                            iTopShapeID = currShapeID;
                    }
                }
            }

            Size size = xControlTriangleShape.getSize();
            Point pos = xControlTriangleShape.getPosition();

            Point a = new Point();
            Point b = new Point();
            Point c = new Point();
            Point d = new Point();

            for(int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")){
                    currShapeID = getController().getNumberOfShape(currShapeName);

                    a.X = pos.X + size.Width/2 - size.Width/iTopShapeID*currShapeID/2;
                    a.Y = pos.Y + size.Height/iTopShapeID*currShapeID;
                    b.X = pos.X + size.Width/2 + size.Width/iTopShapeID*currShapeID/2;
                    b.Y = pos.Y + size.Height/iTopShapeID*currShapeID;
                    c.X = pos.X + size.Width/2 - size.Width/iTopShapeID*(currShapeID-1)/2;
                    c.Y = pos.Y + size.Height/iTopShapeID*(currShapeID-1);
                    d.X = pos.X + size.Width/2 + size.Width/iTopShapeID*(currShapeID-1)/2;
                    d.Y = pos.Y + size.Height/iTopShapeID*(currShapeID-1);

                    setTrapezeShape(xCurrShape, a, b, c, d);
                }
            }
            
        }catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setTrapezeShape(XShape xShape, Point a, Point b, Point c, Point d){

        try {

            Point[] points1 = new Point[4];
            points1[0] = new Point(a.X, a.Y);
            points1[1] = new Point(c.X, c.Y);
            points1[2] = new Point(d.X, d.Y);
            points1[3] = new Point(b.X, b.Y);

            Point[] points2 = new Point[2];
            points2[0] = new Point(a.X, a.Y);
            points2[1] = new Point(b.X, b.Y);

            Point[][] allPoints = new Point[2][];
            allPoints[0] = points1;
            allPoints[1] = points2;

            XPropertySet xPolygonShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPolygonShapeProps.setPropertyValue("PolyPolygon", allPoints);
            
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
                //xShapeProps.setPropertyValue("FillColor", new Integer(0xFF0000));

                xShapeProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
                xShapeProps.setPropertyValue("FillTransparence", new Integer(1000));
                xShapeProps.setPropertyValue("LineStyle", LineStyle.NONE);

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
    public void refreshShapeProperties(){

        XShapes xShapes = getController().getSelectedShapes();

        if (xShapes != null){

            try {

                if(!getGui().isEnableControlDialogImageColor())
                    getGui().enableControlDialogImageColor();


                if( m_Style == USER_DEFINE ){
                    // need to memorize members, if user exit into propsDialog
                    boolean isSelectedAllShape = m_IsSelectAllShape;
                    boolean isGradients = m_IsGradients;
                    boolean isBaseColors = m_IsBaseColors;
                    boolean isBaseColorsWithGradients = m_IsBaseColorsWithGradients;
                    boolean isMonographic = m_IsMonographic;

                    //default values
                    m_IsSelectAllShape = true;
                    m_IsGradients = false;
                    m_IsBaseColors = false;
                    m_IsBaseColorsWithGradients = false;
                    m_IsMonographic = true;

                    m_IsAction = false;

                    getGui().showDiagramPropsDialog();

                    if(m_IsAction){
                   
                        if(m_IsSelectAllShape){
                            setAllShapeProperties();
                        }else{

                            m_IsBaseColors = false;
                            m_IsBaseColorsWithGradients = false;
                            m_IsMonographic = isMonographic;

                            XShape xCurrShape = null;
                            String currShapeName = "";

                            for(int i = 0; i < xShapes.getCount(); i++){
                                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                                currShapeName = getShapeName(xCurrShape);
                                if(currShapeName.endsWith("GroupShape") || currShapeName.contains("PolyPolygonShape0"))
                                    setAllShapeProperties();
                                else
                                    if (currShapeName.contains("PolyPolygonShape")){
                                        setShapeProperties(xCurrShape);
                            }
                    }
                        }
                    }else{
                        m_IsSelectAllShape = isSelectedAllShape;
                        m_IsGradients = isGradients;
                        m_IsBaseColors = isBaseColors;
                        m_IsBaseColorsWithGradients = isBaseColorsWithGradients;
                        m_IsMonographic = isMonographic;
                    }
                    m_IsAction = false;

                }else{

                    XShape xCurrShape = null;
                    String currShapeName = "";

                    for(int i = 0; i < xShapes.getCount(); i++){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        if(currShapeName.endsWith("GroupShape") || currShapeName.contains("PolyPolygonShape0"))
                            setAllShapeProperties();
                        else
                            if (currShapeName.contains("PolyPolygonShape")){
                                if(m_Style == BASE_COLORS || m_Style == BASE_COLORS_GRADIENTS)
                                    setAllShapeProperties();
                                else
                                    setShapeProperties(xCurrShape);
                            }
                    }
                }
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }

    }

    public void setAllShapeProperties(){

        if(m_xShapes != null){

            try {

                XShape xCurrShape = null;
                String currShapeName = "";
                int currShapeID = 0;
                int iTopID = -1;

                for(int i=0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("PolyPolygonShape") && !currShapeName.contains("PolyPolygonShape0")) {
                        if( m_Style == BASE_COLORS || m_Style == BASE_COLORS_GRADIENTS  || (m_Style == USER_DEFINE && m_IsBaseColors)){
                            currShapeID = getController().getNumberOfShape(currShapeName);
                            if(currShapeID > iTopID)
                                iTopID = currShapeID;
                            if(getGui()!= null)
                                getGui().setImageColorOfControlDialog(aCOLORS[(currShapeID - 1) % 8]);
                        }
                        setShapeProperties(xCurrShape);
                    }
                }
                if(getGui()!= null && (m_Style == BASE_COLORS || m_Style == BASE_COLORS_GRADIENTS || (m_Style == USER_DEFINE && m_IsBaseColors) ))
                    getGui().setImageColorOfControlDialog(aCOLORS[iTopID % 8]);
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void setShapeProperties(XShape xShape) {
        setShapeProperties(xShape, "PolyPolygonShape");
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {

        int color = -1;
        if(getGui() != null)
            color = getGui().getImageColorOfControlDialog();
        if(color < 0)
            color = COLOR;

        XPropertySet xProp = null;

        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if( m_Style == DEFAULT ){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                xProp.setPropertyValue("FillColor", new Integer(color));
            }else if( m_Style == GRADIENTS){
                setGradientOfShape(xShape, color, 0xFFFFFF);
            }else if( m_Style == BASE_COLORS){
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                xProp.setPropertyValue("FillColor", new Integer(color));
            } else if( m_Style == BASE_COLORS_GRADIENTS){
                setGradientOfShape(xShape, color, 0xFFFFFF);


            }else if( m_Style == USER_DEFINE){
                if(!m_IsGradients){

                    if(m_IsBaseColors){
                        if(m_IsBaseColorsWithGradients){
                            setGradientOfShape(xShape, color, 0xFFFFFF);
                        }else{
                            xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                            xProp.setPropertyValue("FillColor", new Integer(color));
                        }
                    }else{
                        xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                        xProp.setPropertyValue("FillColor", new Integer(m_iColor));
                        getGui().setImageColorOfControlDialog(m_iColor);
                    }
                }else{
                    setGradientOfShape(xShape, m_iStartColor, m_iEndColor);
                    getGui().disableControlDialogImageColor();
                }
  
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

    public int getColorOfShape(XShape xShape){
        try {
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            return AnyConverter.toInt(xShapeProps.getPropertyValue("FillColor"));
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return -1;
    }


    public void setTextAndTextFitToSize(XShape xShape){
        try {
            //XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            //xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.AUTOFIT);
            XText xText = (XText) UnoRuntime.queryInterface(XText.class, xShape);
            xText.setString(getGui().getDialogPropertyValue( "ControlDialog1", "ControlDialog1.Text.Label" ) );
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

}
