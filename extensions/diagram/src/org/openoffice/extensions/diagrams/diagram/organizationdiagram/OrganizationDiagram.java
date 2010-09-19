package org.openoffice.extensions.diagrams.diagram.organizationdiagram;

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



public class OrganizationDiagram extends Diagram{


    private DiagramTree     m_DiagramTree   = null;

    public static final short   UNDERLING   = 0;
    public static final short   ASSOCIATE   = 1;

    //item hierarhich tpye in diagram
    private short           m_sNewItemHType = UNDERLING;

    // rates of measure of groupShape (e.g.: 10:6)
    protected final int     GROUPWIDTH          = 10;
    protected final int     GROUPHEIGHT         = 6;

    // rates of measure of rectangles
    //(e.g.: WIDTH:HORSPACE 6:1, HEIGHT:VERSPACE 4:3)
    protected final int     WIDTH               = 4;
    protected final int     HORSPACE            = 1;
    protected final int     HEIGHT              = 4;
    protected final int     VERSPACE            = 3;

    protected final short   DEFAULT             = 0;
    protected final short   NOT_MONOGRAPHIC     = 1;
    protected final short   ROUNDED             = 2;
    protected final short   GRADIENTS           = 3;
    protected final short   USER_DEFINE         = 4;

    protected boolean       m_IsGradientAction  = false;
    public static int       _startColor         = 16711680;
    public static int       _endColor           = 8388608;

    
    public OrganizationDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_sNewItemHType = UNDERLING;
    }


    public void setNewItemHType(short n){
        m_sNewItemHType = n;
    }

    public void setGradientAction(boolean bool){
        m_IsGradientAction  = bool;
    }
    @Override
    public String getDiagramType(){
        return "OrganizationDiagram";
    }
    
    public XShapes getShapeGroup(){
        return m_xShapes;
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(4);
    }

    public void createDiagram(int n) {
        try {

            int maxItem = 0;
            if(n>2)
                maxItem = n - 2;   // 0, 1, 2, ...

            int orignGSWidth, horUnit, horSpace, shapeWidth, verUnit, verSpace, shapeHeight, xCoord, yCoord;
            
            orignGSWidth = m_GroupSizeWidth;
            if( (m_GroupSizeWidth / GROUPWIDTH ) <= ( m_GroupSizeHeight / GROUPHEIGHT ) )
                m_GroupSizeHeight = m_GroupSizeWidth * GROUPHEIGHT / GROUPWIDTH;
            else
                m_GroupSizeWidth = m_GroupSizeHeight * GROUPWIDTH / GROUPHEIGHT;

            // set new size of m_xGroupShape for Organigram          
            m_xGroupShape.setSize(new Size(m_GroupSizeWidth,m_GroupSizeHeight));

            int halfDiff = 0;
            if(orignGSWidth > m_GroupSizeWidth)
                halfDiff = (orignGSWidth - m_GroupSizeWidth) / 2;
                m_xGroupShape.setPosition( new Point( m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderTop ) );
            // base shape
            XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft, m_PageProps.BorderBottom, m_GroupSizeWidth, m_GroupSizeHeight);
            m_xShapes.add(xBaseShape);
            XPropertySet xBaseProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xBaseShape);
            xBaseProp.setPropertyValue("FillColor", new Integer(0xFFFF00));
            xBaseProp.setPropertyValue("FillTransparence", new Integer(100));
            xBaseProp.setPropertyValue("LineColor", new Integer(0xFFFFFF));
            xBaseProp.setPropertyValue("LineTransparence", new Integer(1000));
            xBaseProp.setPropertyValue("MoveProtect", new Boolean(true));

            horUnit = m_GroupSizeWidth / ( (maxItem + 1) * WIDTH + maxItem*HORSPACE );
            horSpace = horUnit * HORSPACE;
            shapeWidth = horUnit * WIDTH;

            if(n>1){
                verUnit = m_GroupSizeHeight / ( 2 * HEIGHT + VERSPACE);
                verSpace = verUnit * VERSPACE;
                shapeHeight = verUnit * HEIGHT;
            }else{
                verSpace = 0;
                shapeHeight = m_GroupSizeHeight;
            }

            xCoord = m_PageProps.BorderLeft + maxItem * (shapeWidth + horSpace) / 2;
            yCoord = m_PageProps.BorderTop;

            XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
            m_xShapes.add(xStartShape);
            setMoveProtectOfShape(xStartShape);
 
            setTextFitToSize(xStartShape);
            setShapeProperties(xStartShape, "RectangleShape");
            setColorOfShape(xStartShape, COLOR);
            XShape xRectShape = null;
            for( int i = 2; i <= n; i++ ){
                xRectShape = createShape("RectangleShape", i, m_PageProps.BorderLeft + (i-2) * (shapeWidth + horSpace), yCoord + shapeHeight + verSpace, shapeWidth, shapeHeight);
                m_xShapes.add(xRectShape);
                setMoveProtectOfShape(xRectShape);
                setTextFitToSize(xRectShape);
                setShapeProperties(xRectShape, "RectangleShape");
                setColorOfShape(xRectShape, COLOR);
                XShape xConnectorShape = createShape("ConnectorShape", i);
                m_xShapes.add(xConnectorShape);
                setMoveProtectOfShape(xConnectorShape);
                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectShape, new Integer(0));
            }
            getController().setSelectedShape((Object)xStartShape);
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        }

    }

    @Override
    public void initDiagram() {
        //initial members: m_xDrawPage, m_DiagramID, m_xShapes
        super.initDiagram(); 
        if(m_DiagramTree == null)
            m_DiagramTree = new DiagramTree(this);
        setListsOfDiagramTree();
        m_DiagramTree.setTree(); 
    }

    public boolean searchErrorInTree(){
        if(m_DiagramTree != null){
            setListsOfDiagramTree();
            if((m_DiagramTree.getRectangleListSize() - 1) != m_DiagramTree.getConnectorListSize())
                return true;
            if(!m_DiagramTree.isValidConnectors())
                return true;
            if(m_DiagramTree.setRootItem() != 1)
                return true;
        }
        return false;
    }

    public void repairDiagram(){
        m_DiagramTree.repairTree();
        initDiagram();
    }

    public void setListsOfDiagramTree(){
        try {
            m_DiagramTree.clear();
            XShape xCurrShape = null;
            String currShapeName = "";

            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")) {
                    if (currShapeName.endsWith("RectangleShape0")) {
                        m_DiagramTree.setControlShape(xCurrShape);
                    }else{
                        m_DiagramTree.addToRectangles(xCurrShape);
                    }
                }
                if (currShapeName.contains("ConnectorShape"))
                    m_DiagramTree.addToConnectors(xCurrShape);
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in OrganizationDiagram.setListsOfDigramTree(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.setListsOfDigramTree(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    @Override
    public void refreshDiagram(){
        m_DiagramTree.refresh();
    }

    @Override
    public void addShape(){
        try {
            if(m_DiagramTree != null){
                XShape xSelectedShape = getController().getSelectedShape();
                if(xSelectedShape != null){
                    XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
                    String selectedShapeName = xNamed.getName();

                    int iTopShapeID = -1;
                    XShape xCurrShape = null;
                    String currShapeName = "";

                    if(selectedShapeName.contains("RectangleShape")){// && !selectedShapeName.contains("RectangleShape0")){
                        TreeItem selectedItem = m_DiagramTree.getTreeItem(xSelectedShape);

                        // can't be associate of root item
                        if(selectedItem.getDad() == null && m_sNewItemHType == 1 ) {
                            String title = getGui().getDialogPropertyValue("Strings", "ItemAddError.Title");
                            String message = getGui().getDialogPropertyValue("Strings", "ItemAddError.Message");
                            getGui().showMessageBox(title, message);
                        }else{
                            // adjust iTopShapeID
                            for( int i=0; i < m_xShapes.getCount(); i++ ){
                                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                                currShapeName = getShapeName(xCurrShape);
                                if (currShapeName.contains("RectangleShape")) {
                                    int shapeID = getController().getNumberOfShape(currShapeName);
                                    if (shapeID > iTopShapeID)
                                        iTopShapeID = shapeID;
                                }
                            }

                            if( iTopShapeID <= 0 ){
                                clearEmptyDiagramAndReCreate();
                            }else{

                                iTopShapeID++;
                                XShape xRectangleShape = createShape("RectangleShape", iTopShapeID);
                                
                                m_DiagramTree.addToRectangles(xRectangleShape);

                                //int width = m_DiagramTree.getControlShape().getSize().Width;
                                int width = xSelectedShape.getSize().Width;
                                //int height = m_DiagramTree.getControlShape().getSize().Height;
                                int height = xSelectedShape.getSize().Height;
                                xRectangleShape.setSize(new Size(width, height));

                                int x = 0;
                                int y = 0;
                                short level = -1;
                                TreeItem newTreeItem = null;
                                TreeItem dadItem = null;

                                if(m_sNewItemHType == UNDERLING){
                                    dadItem = selectedItem;
                                    level = (short)(selectedItem.getLevel() + 1);
                                    XShape xPreviousChild = m_DiagramTree.getLastChildShape(xSelectedShape);
                                    if(xPreviousChild != null){
                                        if(level > 2){
                                            x = xPreviousChild.getPosition().X;
                                            y = xPreviousChild.getPosition().Y + (int)(height * 1.5);
                                        }else{
                                            x = xPreviousChild.getPosition().X + 100;
                                            y = xPreviousChild.getPosition().Y ;
                                        }

                                    }else{
                                        x = selectedItem.getPosition().X + 100;
                                        y = selectedItem.getPosition().Y + (int)(height * 1.5);
                                    }

                                    newTreeItem = new TreeItem(m_DiagramTree, xRectangleShape, dadItem, level , (short)0);
                                    if(dadItem.getFirstChild() == null){
                                        dadItem.setFirstChild(newTreeItem);
                                    }else{
                                        TreeItem previousItem = m_DiagramTree.getTreeItem(xPreviousChild);
                                        previousItem.setFirstSibling(newTreeItem);
                                    }

                                }else if(m_sNewItemHType == ASSOCIATE){
                                    dadItem = selectedItem.getDad();
                                    level = selectedItem.getLevel();
                                    if(level > 2){
                                        x = selectedItem.getPosition().X;
                                        y = selectedItem.getPosition().Y + (int)(height * 1.5);
                                     }else{
                                        x = selectedItem.getPosition().X + 100;
                                        y = selectedItem.getPosition().Y;
                                    }

                                    newTreeItem = new TreeItem(m_DiagramTree, xRectangleShape, dadItem, level , (short)0);
                                    if(selectedItem.getFirstSibling() != null)
                                        newTreeItem.setFirstSibling(selectedItem.getFirstSibling());
                                    selectedItem.setFirstSibling(newTreeItem);
                                }

                                xRectangleShape.setPosition(new Point(x,y));
                                m_xShapes.add(xRectangleShape);
                                setMoveProtectOfShape(xRectangleShape);
                                setTextFitToSize(xRectangleShape);
                                setShapeProperties(xRectangleShape, "RectangleShape");
                                int color = getGui().getImageColorOfControlDialog();
                                if(color < 0)
                                    color = COLOR;
                                setColorOfShape(xRectangleShape, color);

                                if(iTopShapeID > 1){
                                    // set connector shape
                                    XShape xConnectorShape = createShape("ConnectorShape", iTopShapeID);
                                    
                                    m_DiagramTree.addToConnectors(xConnectorShape);

                                    m_xShapes.add(xConnectorShape);
                                    setMoveProtectOfShape(xConnectorShape);
                                    XShape xStartShape = null;
                                    Integer endShapeConnPos = new Integer(0);
                                    if(m_sNewItemHType == UNDERLING){
                                        xStartShape = getController().getSelectedShape();
                                        if(level > 2)
                                           endShapeConnPos = new Integer(3);
                                    }else if(m_sNewItemHType == ASSOCIATE){
                                        xStartShape = selectedItem.getDad().getRectangleShape();
                                        if(level > 2)
                                           endShapeConnPos = new Integer(3);
                                    }
                                    setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectangleShape, endShapeConnPos);
                                }
                            }
                        }
                    }else if(selectedShapeName.contains("GroupShape")){
                        for( int i=0; i < m_xShapes.getCount(); i++ ){
                            xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                            currShapeName = getShapeName(xCurrShape);
                            if (currShapeName.contains("RectangleShape")) {
                                int shapeID = getController().getNumberOfShape(currShapeName);
                                if (shapeID > iTopShapeID)
                                    iTopShapeID = shapeID;
                            }
                        }
                        if( iTopShapeID <= 0 )
                            clearEmptyDiagramAndReCreate();
                    }  // else if
                }
            }
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        } catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.createDiagram(). Message:\n" + ex.getLocalizedMessage());
        }
    }

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
            m_xDrawPage.remove(m_xGroupShape);
            createDiagram(1);
        } catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in OrganizationDiagram.clearEmptyDiagramAndReCreate(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.clearEmptyDiagramAndReCreate(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    @Override
    public void removeShape(){
        XShape xSelectedShape = getController().getSelectedShape();
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
                    TreeItem selectedItem = m_DiagramTree.getTreeItem(xSelectedShape);
                        
                    TreeItem dadItem = selectedItem.getDad();
                    if(selectedItem.equals(dadItem.getFirstChild())){
                        if(selectedItem.getFirstSibling() != null){
                            dadItem.setFirstChild(selectedItem.getFirstSibling());                               
                        }else{
                            dadItem.setFirstChild(null);
                        }
                    }else{
                        TreeItem previousSibling = m_DiagramTree.getPreviousSibling(selectedItem);
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
                        
                    XShape xConnShape = m_DiagramTree.getDadConnectorShape(xSelectedShape);
                    if (xConnShape != null){
                        m_DiagramTree.removeFromConnectors(xConnShape);
                        m_xShapes.remove(xConnShape);
                    }
                    m_DiagramTree.removeFromRectangles(xSelectedShape);
                    m_xShapes.remove(xSelectedShape);
                    cutSelectedItem(selectedItem);
                    getController().setSelectedShape((Object)xDadShape);

                }
            }
        }
    }

    public void removeShapeFromGroup(XShape xShape){
        m_xShapes.remove(xShape);
    }

    public void cutSelectedItem(TreeItem item){
        item.setDad(null);
        item.setFirstChild(null);
        item.setFirstSibling(null);
    }

    public void selectGroupShape(){
        if(m_xShapes != null)
            getController().setSelectedShape((Object)m_xShapes);
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
            System.err.println("UnknownPropertyException in VennDiagram.refreshShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in VennDiagram.refreshShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in VennDiagram.refreshShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (IndexOutOfBoundsException ex) {
            System.err.println("IndexOutOfBoundsException in VennDiagram.refreshShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in VennDiagram.refreshShapeProperties(). Message:\n" + ex.getLocalizedMessage());
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
            System.err.println("IndexOutOfBoundsException in VennDiagram.setAllShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in VennDiagram.setAllShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {

        int color = getGui().getImageColorOfControlDialog();
        if(color < 0)
            color = COLOR;
        XPropertySet xProp = null;

        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            //if( m_Style != GRADIENTS ){
            if( m_Style != GRADIENTS || !( m_Style == USER_DEFINE && m_IsGradients == true)){
                if(!getGui().isEnableControlDialogImageColor())
                    getGui().enableControlDialogImageColor();
            }
            if( m_Style == DEFAULT ){
                xProp.setPropertyValue("FillColor", new Integer(color));
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("CornerRadius", new Integer(0));
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
            }else if( m_Style == NOT_MONOGRAPHIC){
                xProp.setPropertyValue("FillColor", new Integer(color));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                xProp.setPropertyValue("CornerRadius", new Integer(0));
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
            }else if( m_Style == ROUNDED){
                xProp.setPropertyValue("FillColor", new Integer(color));
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("CornerRadius", new Integer(600));
                xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
            } else if( m_Style == GRADIENTS){
                getGui().disableControlDialogImageColor();
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("CornerRadius", new Integer(0));
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
            System.err.println("IllegalArgumentException in VennDiagram.setShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in VennDiagram.setShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in VennDiagram.setShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in VennDiagram.setShapeProperties(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public void setMoveProtectOfShape(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (Exception ex) {
            System.err.println("Exception in OrganizationDiagram.setMoveProtectOfShape(). Message:\n" + ex.getLocalizedMessage());
        }

    }

    public void setTextFitToSize(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
        } catch (Exception ex) {
            System.err.println("Exception in OrganizationDiagram.setTextFitToSize(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xPropText.setPropertyValue("StartShape", xStartShape);
            xPropText.setPropertyValue("EndShape", xEndShape);
            xPropText.setPropertyValue("StartGluePointIndex", startIndex);
            xPropText.setPropertyValue("EndGluePointIndex", endIndex);
            xPropText.setPropertyValue("LineWidth",new Integer(100));
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in OrganizationDiagram.setTextOfRectangle(). Message:\n" + ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in OrganizationDiagram.setTextOfRectangle(). Message:\n" + ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in OrganizationDiagram.setTextOfRectangle(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.setTextOfRectangle(). Message:\n" + ex.getLocalizedMessage());
        }
    }

}
