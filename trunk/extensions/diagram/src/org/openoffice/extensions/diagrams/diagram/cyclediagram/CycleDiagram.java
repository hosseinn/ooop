package org.openoffice.extensions.diagrams.diagram.cyclediagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.PolygonFlags;
import com.sun.star.drawing.TextFitToSizeType;
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
import org.openoffice.extensions.diagrams.Controller;
import org.openoffice.extensions.diagrams.Gui;
import org.openoffice.extensions.diagrams.diagram.Diagram;


public class CycleDiagram extends Diagram {


    private int             m_GroupSize     = 0;

    protected final short   DEFAULT         = 0;
    protected final short   NOT_MONOGRAPHIC = 1;
    protected final short   BASE_COLORS     = 2;
    protected final short   WITH_FRAME      = 3;
    protected final short   USER_DEFINE     = 4;

    private static short    _count          = 0;


    public CycleDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
    }

    @Override
    public String getDiagramTypeName(){
        return "CycleDiagram";
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(2);
    }

    public void createDiagram(int n){
        try {
            if(m_xDrawPage != null && m_xShapes != null){

                m_GroupSize = m_DrawAreaWidth <= m_DrawAreaHeight ? m_DrawAreaWidth : m_DrawAreaHeight;
                m_xGroupShape.setSize( new Size( m_GroupSize, m_GroupSize ) );

                int halfDiff = 0;
                if(m_GroupSize < m_DrawAreaWidth)
                    halfDiff = (m_DrawAreaWidth - m_GroupSize) / 2;
                m_xGroupShape.setPosition( new Point( m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderTop ) );

                int controlEllipseSize = m_GroupSize / 4 * 3;
                int xR;
                int yR;
                xR = yR = controlEllipseSize / 2;
                Point middlePoint = new Point(m_GroupSize / 2 + m_PageProps.BorderLeft + halfDiff, m_GroupSize / 2 + m_PageProps.BorderTop);
                //ControlEllipse
                Point coord = new Point(middlePoint.X - xR, middlePoint.Y - yR);
                Size size = new Size(controlEllipseSize, controlEllipseSize);
                XShape xControlShape = createShape("EllipseShape", 0, coord.X, coord.Y, size.Width, size.Height);
                m_xShapes.add(xControlShape);
                setInvisibleFeatures(xControlShape);

                XShape xBezierShape = null;
                XShape xRectangleShape = null;

                for(int i = 1; i <= n; i++){
                    xRectangleShape = createShape("RectangleShape", i);
                    m_xShapes.add(xRectangleShape);
                    setInvisibleFeatures(xRectangleShape);
                    setTextAndTextFitToSize(xRectangleShape);
                    xBezierShape = createShape("ClosedBezierShape", i);
                    m_xShapes.add(xBezierShape);
                    setColorOfShape(xBezierShape, COLOR);
                }
                
                refreshDiagram();
                getController().setSelectedShape((Object)m_xShapes);
                if(getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(COLOR);
            } 
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void initDiagram(){

        super.initDiagram();
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeNum;
        int iTopShapeID = 0;
        XShape xTopShape = null;

        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("ClosedBezierShape")) {
                    currShapeNum = getController().getNumberOfShape(currShapeName);
                    if (currShapeNum > iTopShapeID){
                        iTopShapeID = currShapeNum;
                        xTopShape = xCurrShape;
                    }
                }
            }

            int color = -1;
            if(m_Style == BASE_COLORS)
                color = aCOLORS[iTopShapeID % 8];
            else
                if(xCurrShape != null)
                    color = getColorOfShape(xTopShape);
            
            if( color >= 0 && getGui() != null)
                getGui().setImageColorOfControlDialog(color);

        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void addShape(){

        XShape xCurrShape = null;
        String currShapeName = "";
        ArrayList<XShape> rectangleShapeList;
        ArrayList<XShape> bezierShapeList;


        try {

            bezierShapeList = new ArrayList();
            rectangleShapeList = new ArrayList();

            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape"))
                    rectangleShapeList.add(xCurrShape);
                if (currShapeName.contains("ClosedBezierShape"))
                    bezierShapeList.add(xCurrShape);
            }

            XShape xSelectedShape = getController().getSelectedShape();
            String selectedShapeName = getShapeName(xSelectedShape);

            int newBezierShapeID;
            int newRectShapeID;
            int iTopShapeID = 0;
            int shapeID;

            //adjust the iTopShapeID
            for(XShape currRectXShape : rectangleShapeList){
                shapeID = getController().getNumberOfShape(getShapeName(currRectXShape));
                if(shapeID > iTopShapeID)
                    iTopShapeID = shapeID;
            }

            if(selectedShapeName.endsWith("EllipseShape0") || selectedShapeName.endsWith("GroupShape")){    
                newRectShapeID = newBezierShapeID = iTopShapeID + 1;
            }else{
                int selectedShapeID = getController().getNumberOfShape(selectedShapeName);
                newRectShapeID = selectedShapeID;
                if(selectedShapeName.contains("RectangleShape"))
                    newRectShapeID = selectedShapeID + 1;
                increaseShapeIDs(newRectShapeID, "RectangleShape");
                newBezierShapeID = selectedShapeID + 1;
                increaseShapeIDs(newBezierShapeID, "ClosedBezierShape");
            }

            int color = -1;
            if(getGui() != null)
                color = getGui().getImageColorOfControlDialog();
            if(color < 0)
                color = COLOR;

            XShape xRectangleshape = createShape("RectangleShape", newRectShapeID);
            m_xShapes.add(xRectangleshape);
            setInvisibleFeatures(xRectangleshape);
            setTextAndTextFitToSize(xRectangleshape);
            setShapeProperties(xRectangleshape, "RectangleShape");
            XShape xBezierShape = createShape("ClosedBezierShape", newBezierShapeID);
            m_xShapes.add(xBezierShape);
            setShapeProperties(xBezierShape, "ClosedBezierShape");
            setColorOfShape(xBezierShape, color);


            if(m_Style == BASE_COLORS || (m_Style == USER_DEFINE && m_IsBaseColors)){
                if(selectedShapeName.endsWith("EllipseShape0") || selectedShapeName.endsWith("GroupShape")){
                    if(getGui()!= null)
                        getGui().setImageColorOfControlDialog(aCOLORS[newBezierShapeID  % 8]);
                }else{
                    if(getGui()!= null){

                        _count++;
                        if(_count == 9)
                            _count = 1;
                        getGui().setImageColorOfControlDialog(aCOLORS[(newBezierShapeID + _count)  % 8]);

                        // random colors
                        //int selectedShapeColor = getColorOfShapeByID(getController().getNumberOfShape(selectedShapeName));
                        //int newBezierShapeColor = getColorOfShapeByID(newBezierShapeID);
                        
                        //int newColor; // new color between selectedShape and newBezierShape
                        //do{
                        //   newColor = aCOLORS[((int)(Math.random()*8)) % 8];
                        //}while(newColor == selectedShapeColor || newColor == newBezierShapeColor);
                        
                        //getGui().setImageColorOfControlDialog(newColor);
                    }
                }
            }


            if(iTopShapeID < 1){

                color = -1;
                if(getGui() != null)
                    color = getGui().getImageColorOfControlDialog();
                if(color < 0)
                    color = COLOR;

                xRectangleshape = createShape("RectangleShape", newRectShapeID + 1);
                m_xShapes.add(xRectangleshape);
                setInvisibleFeatures(xRectangleshape);
                setTextAndTextFitToSize(xRectangleshape);
                setShapeProperties(xRectangleshape, "RectangleShape");
                xBezierShape = createShape("ClosedBezierShape", newBezierShapeID + 1);
                m_xShapes.add(xBezierShape);
                setShapeProperties(xBezierShape, "ClosedBezierShape");
                setColorOfShape(xBezierShape, color);

                if(m_Style == BASE_COLORS)
                    if(getGui()!= null)
                        getGui().setImageColorOfControlDialog(aCOLORS[(newBezierShapeID + 1)  % 8]);
            }

        } catch (IndexOutOfBoundsException ex) {
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
            for( int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("ClosedBezierShape")) {
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
 
    @Override
    public void removeShape() {

        XShape xSelectedShape = getController().getSelectedShape();
        String selectedShapeName = getShapeName(xSelectedShape);
        int iTopShapeID = 0;
        String selectedShapeType = "";
        int selectedShapeID;
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;
        XShape newSelectedXShape = null;
        int newSelectedShapeID = 1;
        

        if( selectedShapeName.contains("RectangleShape") || selectedShapeName.contains("ClosedBezierShape") ){  
            
            try{

                for( int i = 0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeID = getController().getNumberOfShape(getShapeName(xCurrShape));
                        if(currShapeID > iTopShapeID)
                            iTopShapeID = currShapeID;
                }

                if(iTopShapeID <= 2){

                    ArrayList<XShape> rectShapesList = new ArrayList<XShape>();
                    XShape controlEllipseShape = null;
                    for( int i = 0; i < m_xShapes.getCount(); i++ ){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        if(currShapeName.contains("RectangleShape"))
                            rectShapesList.add(xCurrShape);
                        if(currShapeName.contains("EllipseShape"))
                            controlEllipseShape = xCurrShape;
                    }

                    for(XShape xRectShape : rectShapesList)
                        m_xShapes.remove(xRectShape);
                    removeSingleItemsAndShrinkIDs();
                    getController().setSelectedShape((Object)controlEllipseShape);

                }else{

                    selectedShapeID = getController().getNumberOfShape(selectedShapeName);
                    if(selectedShapeID > 1)
                        newSelectedShapeID = selectedShapeID - 1;

                    m_xShapes.remove(xSelectedShape);
                    removeSingleItemsAndShrinkIDs();

                    if(selectedShapeName.contains("RectangleShape"))
                        selectedShapeType = "RectangleShape";
                    if(selectedShapeName.contains("ClosedBezierShape"))
                        selectedShapeType = "ClosedBezierShape";

                    for( int i = 0; i < m_xShapes.getCount(); i++ ){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        if(currShapeName.contains(selectedShapeType)){
                            currShapeID = getController().getNumberOfShape(currShapeName);
                            if(currShapeID == newSelectedShapeID)
                                newSelectedXShape = xCurrShape;
                        }
                    }

                    if(newSelectedXShape != null)
                        getController().setSelectedShape((Object)newSelectedXShape);
                    else
                        getController().setSelectedShape((Object)m_xShapes);
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public void increaseShapeIDs(int itemID){
        increaseShapeIDs(itemID, "");
    }

    public void increaseShapeIDs(int itemID, String shapeType){

        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;

        try{
            for( int i = 0; i < m_xShapes.getCount(); i++ ){

                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                currShapeID = getController().getNumberOfShape(currShapeName);
                
                if( (shapeType.equals("RectangleShape")  || shapeType.equals("")) && currShapeName.contains("RectangleShape")){
                    if(currShapeID >= itemID){
                        String newShapeName = "";
                        newShapeName = currShapeName.split("RectangleShape", 2)[0] + "RectangleShape" + (currShapeID + 1);
                        XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xCurrShape );
                        xNamed.setName(newShapeName);
                    }
                }
                if( (shapeType.equals("ClosedBezierShape") || shapeType.equals("")) && currShapeName.contains("ClosedBezierShape")){
                    if(currShapeID >= itemID){
                        String newShapeName = "";
                        newShapeName = currShapeName.split("ClosedBezierShape", 2)[0] + "ClosedBezierShape" + (currShapeID + 1);
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

    public void decreaseShapeIDs(int itemID){
        decreaseShapeIDs(itemID, "");
    }

    public void decreaseShapeIDs(int itemID, String shapeType){

        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;

        try{
            for( int i = 0; i < m_xShapes.getCount(); i++ ){

                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                currShapeID = getController().getNumberOfShape(currShapeName);

                if( (shapeType.equals("RectangleShape")  || shapeType.equals("")) && currShapeName.contains("RectangleShape")){
                    if(currShapeID > itemID){
                        String newShapeName = "";
                        newShapeName = currShapeName.split("RectangleShape", 2)[0] + "RectangleShape" + (currShapeID - 1);
                        XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xCurrShape );
                        xNamed.setName(newShapeName);
                    }
                }
                if( (shapeType.equals("ClosedBezierShape") || shapeType.equals("")) && currShapeName.contains("ClosedBezierShape")){
                    if(currShapeID > itemID){
                        String newShapeName = "";
                        newShapeName = currShapeName.split("ClosedBezierShape", 2)[0] + "ClosedBezierShape" + (currShapeID - 1);
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

    @Override
    public void refreshDiagram(){
        
        XShape xCurrShape = null;
        String currShapeName = "";
        XShape xControlEllipseShape = null;
        ArrayList<XShape> rectangleShapeList = new ArrayList();
        ArrayList<XShape> bezierShapeList = new ArrayList();

        try {
            
            removeSingleItemsAndShrinkIDs();

            //fill the arrayLists and adjust the controrEllipse
            for(int i = 0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("EllipseShape0")) {
                    xControlEllipseShape = xCurrShape;
                }else{
                    if (currShapeName.contains("RectangleShape"))
                        rectangleShapeList.add(xCurrShape);
                    if (currShapeName.contains("ClosedBezierShape"))
                        bezierShapeList.add(xCurrShape);
                }
            }

            int iTopShapeID = 0;
            int shapeID;

            //adjust the iTopShapeID
            for(XShape currRectXShape : rectangleShapeList){
                shapeID = getController().getNumberOfShape(getShapeName(currRectXShape));
                if(shapeID > iTopShapeID)
                    iTopShapeID = shapeID;
            }

            int numOfItems = iTopShapeID;

            Point point = xControlEllipseShape.getPosition();
            Size size = xControlEllipseShape.getSize();
            int xR = size.Width/2;
            int yR = size.Height/2;
            Point middlePoint = new Point( point.X + xR, point.Y + yR);
            int rectShapeWidth = (int)(size.Width/numOfItems*1.8);
            int rectShapeHeight = size.Height/numOfItems;
            Point rectMiddlePoint = new Point( middlePoint.X - rectShapeWidth/2, middlePoint.Y - rectShapeHeight/2);

            //set positions of shapes
            double    angle = - Math.PI/2;
            int radius = size.Width/2 <= size.Height/2 ? size.Width/2 : size.Height/2;

            if(numOfItems > 4 ){
                double newRadius;
                if(numOfItems == 5)
                    newRadius = (numOfItems / 100.0 + 1.0) * radius;
                else
                    newRadius = (numOfItems / 80.0 + 1.0) * radius;
                if(numOfItems > 20 )
                    newRadius = 1.25 * radius;
                radius = (int)newRadius;
            }
            int radius2, radius3, radius4, radius5, radius6;
            radius2 = radius3 = radius4 = radius5 = radius6 = 0;

            double point1Diff, point2Diff, point3Diff, point4Diff, point5Diff, point6Diff, point7Diff, point8Diff, point9Diff;
            point1Diff = point2Diff = point3Diff = point4Diff = point5Diff =
            point6Diff = point7Diff= point8Diff = point9Diff = 0.0;

            if(numOfItems == 2){
                rectShapeWidth = (int)(rectShapeWidth*0.5);
                rectShapeHeight = (int)(rectShapeHeight*0.6);
                rectMiddlePoint = new Point( middlePoint.X - rectShapeWidth/2, middlePoint.Y - rectShapeHeight/2);
                radius2 = (int)(radius + radius/(numOfItems*3));
                radius3 = (int)(radius + radius/(numOfItems*1.4));
                radius4 = (int)(radius - radius/(numOfItems*1.3));
                radius5 = (int)(radius - radius/(numOfItems*3));
                radius6 = (int)(radius);
                double diff = Math.PI/16 + 2 * Math.PI/numOfItems;
                point1Diff = - Math.PI/numOfItems/2 + diff;
                point2Diff = - Math.PI/numOfItems/4 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI/numOfItems/3/20 + diff;
                point5Diff = Math.PI/numOfItems/3 + diff;
                point6Diff = - Math.PI/numOfItems/3/10 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI/numOfItems/4 + diff;
                point9Diff =  - Math.PI/numOfItems/2 + diff;
            }

            if(numOfItems == 3){
                rectShapeWidth = (int)(size.Width/numOfItems*1.35);
                rectShapeHeight = (int)(size.Height/numOfItems*0.85);
                rectMiddlePoint.X =  rectMiddlePoint.X + rectShapeWidth/4;
                rectMiddlePoint.Y =  rectMiddlePoint.Y + rectShapeHeight/4;
                radius2 = (int)(radius + radius/(numOfItems*2.5));
                radius3 = (int)(radius + radius/(numOfItems*1.2));
                radius4 = (int)(radius - radius/(numOfItems*1.2));
                radius5 = (int)(radius - radius/(numOfItems*2.5));
                radius6 = (int)(radius - radius/(numOfItems*10));
                double diff = 2 * Math.PI/numOfItems;
                point1Diff = - Math.PI/numOfItems/2 + diff;
                point2Diff = - Math.PI/numOfItems/4 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI/numOfItems/3/8 + diff;
                point5Diff = Math.PI/numOfItems/3 + diff;
                point6Diff = - Math.PI/numOfItems/3/8 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI/numOfItems/4 + diff;
                point9Diff =  - Math.PI/numOfItems/2 + diff;
            }

            if(numOfItems == 4){
                rectShapeWidth = (int)(size.Width/numOfItems*1.6);
                rectMiddlePoint.X =  rectMiddlePoint.X + rectShapeWidth/16;
            }

            if(numOfItems >= 4){
                radius2 = (int)(radius + radius/(numOfItems*2));
                radius3 = (int)(radius + radius/numOfItems);
                radius4 = (int)(radius - radius/numOfItems);
                radius5 = (int)(radius - radius/(numOfItems*2));
                radius6 = (int)(radius - radius/(numOfItems*10));
                double diff = 2 * Math.PI/numOfItems;
                point1Diff = - Math.PI/numOfItems/3 + diff;
                point2Diff = - Math.PI/numOfItems/3/2 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI/numOfItems/3/8 + diff;
                point5Diff = Math.PI/numOfItems/3 + diff;
                point6Diff = - Math.PI/numOfItems/3/8 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI/numOfItems/3/2 + diff;
                point9Diff =  - Math.PI/numOfItems/3 + diff;
            }

            XShape xBezierShape = null;
            XShape xRectShape = null;
            int xCoord;
            int yCoord;

            //make numCircle(number of circle) angles around the middlePoint and instantate the shapes
            for(int i=0; i < numOfItems; i++, angle += 2.0 * Math.PI / numOfItems){
               
                XShape xCurrBezierShape = null;
                int currShapeID;
                for(int j = 0; j < bezierShapeList.size(); j++){
                    xCurrBezierShape = bezierShapeList.get(j);
                     currShapeID = getController().getNumberOfShape(getShapeName(xCurrBezierShape));
                     if(currShapeID == i+1)
                      xBezierShape = xCurrBezierShape;
                }

                if(xBezierShape != null){
                    Point point1 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point1Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point1Diff)));
                    Point point2 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point2Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point2Diff)));
                    Point point3 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point3Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point3Diff)));
                    Point point4 = new Point((int)(middlePoint.X + radius3 * Math.cos(angle + point4Diff)),(int)(middlePoint.Y + radius3 * Math.sin(angle + point4Diff)));
                    Point point5 = new Point((int)(middlePoint.X + radius6 * Math.cos(angle + point5Diff)),(int)(middlePoint.Y + radius6 * Math.sin(angle + point5Diff)));
                    Point point6 = new Point((int)(middlePoint.X + radius4 * Math.cos(angle + point6Diff)),(int)(middlePoint.Y + radius4 * Math.sin(angle + point6Diff)));
                    Point point7 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point7Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point7Diff)));
                    Point point8 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point8Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point8Diff)));
                    Point point9 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point9Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point9Diff)));
                    setBezierShapeSize(xBezierShape, point1, point2, point3, point4, point5 ,point6 ,point7, point8, point9);
                }

                XShape xCurrRectShape = null;
                for(int j = 0; j < rectangleShapeList.size(); j++){
                    xCurrRectShape = rectangleShapeList.get(j);
                    currShapeID = getController().getNumberOfShape(getShapeName(xCurrRectShape));
                    if(currShapeID == i+1)
                        xRectShape = xCurrRectShape;
                }

                double rectRadius = 0.0;
                if(numOfItems == 2)
                    rectRadius = radius * 0.8;
                else
                    rectRadius = radius;

                xCoord = (int)(rectMiddlePoint.X + rectRadius * Math.cos(angle + Math.PI/numOfItems));
                yCoord = (int)(rectMiddlePoint.Y + rectRadius * Math.sin(angle + Math.PI/numOfItems));
                xRectShape.setPosition(new Point(xCoord, yCoord));
                xRectShape.setSize(new Size(rectShapeWidth,rectShapeHeight));
            }

        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void removeSingleItemsAndShrinkIDs(){

        XShape xShape = null;
        String shapeName = "";
        int shapeID;
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;
        boolean isSingleItem = true;
        ArrayList<XShape> singleItemList;
        int iTopShapeID = 0;

        try {

            singleItemList = new ArrayList<XShape>();
  
            for(int i = 0; i < m_xShapes.getCount(); i++){

                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                shapeName = getShapeName(xShape);

                if ( shapeName.contains("RectangleShape") || shapeName.contains("ClosedBezierShape") ){

                    shapeID = getController().getNumberOfShape(shapeName);

                    if (shapeName.contains("RectangleShape")){
                        isSingleItem = true;
                        for(int j = 0; j < m_xShapes.getCount(); j++){
                            xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                            currShapeName = getShapeName(xCurrShape);
                            if (currShapeName.contains("ClosedBezierShape")){
                                currShapeID = getController().getNumberOfShape(currShapeName);
                                if(currShapeID == shapeID)
                                    isSingleItem = false;
                            }
                        }
                    }
                    if (shapeName.contains("ClosedBezierShape")){
                        isSingleItem = true;
                        for(int j = 0; j < m_xShapes.getCount(); j++){
                            xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                            currShapeName = getShapeName(xCurrShape);
                            if (currShapeName.contains("RectangleShape")){
                                currShapeID = getController().getNumberOfShape(currShapeName);
                                if(currShapeID == shapeID)
                                    isSingleItem = false;
                            }
                        }
                    }

                    if(isSingleItem)
                        singleItemList.add(xShape);
                }
            }

            for(XShape xSingleShape : singleItemList)
                m_xShapes.remove(xSingleShape);

            for(int i = 0; i < m_xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                shapeID = getController().getNumberOfShape(getShapeName(xShape));
                if(shapeID > iTopShapeID)
                    iTopShapeID = shapeID;
            }

            boolean isNotExistingItem;

            for(int i = iTopShapeID; i > 0; i--){
                isNotExistingItem = true;
                for(int j = 0; j < m_xShapes.getCount(); j++){
                    xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    shapeName = getShapeName(xShape);
                    if (shapeName.contains("RectangleShape")){
                        shapeID = getController().getNumberOfShape(shapeName);
                        if(shapeID == i)
                            isNotExistingItem = false;
                    }

                }
                if(isNotExistingItem)
                    decreaseShapeIDs(i);
            }

        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }

    }

    // need to pass bezierShape and its 9 points
    public void setBezierShapeSize(XShape xBezierShape, Point point1, Point point2, Point point3, Point point4, Point point5, Point point6, Point point7, Point point8, Point point9 ){
        try {
            
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
            pointCoords[10] = point9;

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

            XPropertySet xBezierShapeProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xBezierShape);
            xBezierShapeProps.setPropertyValue("PolyPolygonBezier", aCoords);

        }catch (UnknownPropertyException ex) {
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
        try {

            if( m_Style == USER_DEFINE ){
                // need to memorize members, if user exit into propsDialog
                boolean isSelectAllShape = m_IsSelectAllShape;
                boolean isBaseColors = m_IsBaseColors;
                boolean isMonographic = m_IsMonographic;
                boolean isFrame = m_IsFrame;

                //default values
                m_IsSelectAllShape = true;
                m_IsBaseColors = false;
                m_IsMonographic = true;
                m_IsFrame = false;

                m_IsAction = false;

                getGui().showDiagramPropsDialog();

                if(m_IsAction){

                    if(m_IsSelectAllShape){
                        setAllShapeProperties();
                    }else{

                        m_IsBaseColors = false;
                        m_IsMonographic = isMonographic;
                        m_IsFrame = isFrame;

                        XShape xCurrShape = null;
                        String currShapeName = "";
                        XShapes xShapes = getController().getSelectedShapes();

                        if (xShapes != null){
                            for(int i = 0; i < xShapes.getCount(); i++){
                                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                                currShapeName = getShapeName(xCurrShape);
                                if(currShapeName.endsWith("GroupShape") || currShapeName.contains("EllipseShape")){
                                    setAllShapeProperties();
                                }else{
                                    if (currShapeName.contains("RectangleShape") || currShapeName.contains("ClosedBezierShape")){
                                        XShape pairOfXShape = getPairOfShape(xCurrShape);
                                        setShapeProperties(pairOfXShape);
                                        setShapeProperties(xCurrShape);
                                    }
                                }
                            }
                        }
                    }
                }else{
                    m_IsSelectAllShape = isSelectAllShape;
                    m_IsBaseColors = isBaseColors;
                    m_IsMonographic = isMonographic;
                    m_IsFrame = isFrame;
                }
                m_IsAction = false;
            }else{
                
                XShape xCurrShape = null;
                String currShapeName = "";
                XShapes xShapes = getController().getSelectedShapes();

                if (xShapes != null){
                    for(int i = 0; i < xShapes.getCount(); i++){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        if(currShapeName.endsWith("GroupShape") || currShapeName.contains("EllipseShape"))
                            setAllShapeProperties();
                        else
                            if (currShapeName.contains("RectangleShape") || currShapeName.contains("ClosedBezierShape")){
                                if(m_Style == BASE_COLORS)
                                    setAllShapeProperties();
                                else{
                                    XShape pairOfXShape = getPairOfShape(xCurrShape);
                                    setShapeProperties(pairOfXShape);
                                    setShapeProperties(xCurrShape);
                                }
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

    public XShape getPairOfShape(XShape xShape){

        String shapeName = getShapeName(xShape);
        int shapeID = getController().getNumberOfShape(shapeName);
        XShape xCurrShape = null;
        String currShapeName = "";
        int currShapeID;

        try {
            if (shapeName.contains("RectangleShape")){
                for(int i = 0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("ClosedBezierShape")) {
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID == shapeID)
                            return xCurrShape;
                    }
                }
            }
            if (shapeName.contains("ClosedBezierShape")){
                for(int i = 0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("RectangleShape")) {
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if(currShapeID == shapeID)
                            return xCurrShape;
                    }
                }
            }

        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }

    public void setAllShapeProperties(){
        try {
            XShape xCurrShape = null;
            String currShapeName = "";
            int currShapeID = 0;
            int iTopID = -1;

            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("ClosedBezierShape") || currShapeName.contains("RectangleShape")) {
                    if (currShapeName.contains("ClosedBezierShape")){
                        if( m_Style == BASE_COLORS || (m_Style == USER_DEFINE && m_IsBaseColors) ){
                            currShapeID = getController().getNumberOfShape(currShapeName);
                            if(currShapeID > iTopID)
                                iTopID = currShapeID;
                            if(getGui()!= null)
                                getGui().setImageColorOfControlDialog(aCOLORS[(currShapeID - 1) % 8]);
                        }
                    }
                    setShapeProperties(xCurrShape);
                }
            }
            if(getGui()!= null && ( m_Style == BASE_COLORS || (m_Style == USER_DEFINE && m_IsBaseColors)) )
                getGui().setImageColorOfControlDialog(aCOLORS[iTopID % 8]);
         
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setShapeProperties(XShape xShape) {
        if (getShapeName(xShape).contains("RectangleShape"))
            setShapeProperties(xShape, "RectangleShape");
            if (getShapeName(xShape).contains("ClosedBezierShape"))
                setShapeProperties(xShape, "ClosedBezierShape");
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
                if (type.equals("RectangleShape"))
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if (type.equals("ClosedBezierShape")){
                    xProp.setPropertyValue("FillColor", new Integer(color));
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                }
            }else if( m_Style == NOT_MONOGRAPHIC){
                if (type.equals("RectangleShape"))
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if (type.equals("ClosedBezierShape")){
                    xProp.setPropertyValue("FillColor", new Integer(color));
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
            }else if( m_Style == BASE_COLORS){
                if (type.equals("RectangleShape"))
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if (type.equals("ClosedBezierShape")){
                    xProp.setPropertyValue("FillColor", new Integer(color));
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                }
            } else if( m_Style == WITH_FRAME){
                if (type.equals("RectangleShape"))
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                if (type.equals("ClosedBezierShape")){
                    xProp.setPropertyValue("FillColor", new Integer(color));
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                }
            }else if( m_Style == USER_DEFINE){
                if (type.equals("RectangleShape")){
                    if(m_IsFrame)
                        xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    else
                        xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
                if (type.equals("ClosedBezierShape")){
                    if(m_IsBaseColors)
                        xProp.setPropertyValue("FillColor", new Integer(color));
                    else{
                        xProp.setPropertyValue("FillColor", new Integer(m_iColor));
                        getGui().setImageColorOfControlDialog(m_iColor);
                    }

                    if(m_IsMonographic)
                        xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    else
                        xProp.setPropertyValue("LineStyle", LineStyle.NONE);
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

    public void setInvisibleFeatures(XShape xShape){
        try {
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xShapeProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
            xShapeProps.setPropertyValue("FillTransparence", new Integer(1000));
            xShapeProps.setPropertyValue("LineStyle", LineStyle.NONE);
            //xShapeProps.setPropertyValue("LineTransparence", new Integer(1000));
            xShapeProps.setPropertyValue("MoveProtect", new Boolean(true));
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

    public void setTextAndTextFitToSize(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
            XText xText = (XText) UnoRuntime.queryInterface(XText.class, xShape);
            xText.setString(getGui().getDialogPropertyValue( "ControlDialog1", "ControlDialog1.Text.Label" ) );
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

}
