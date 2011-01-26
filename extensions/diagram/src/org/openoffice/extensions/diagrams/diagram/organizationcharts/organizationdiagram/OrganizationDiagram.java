package org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import org.openoffice.extensions.diagrams.Controller;
import org.openoffice.extensions.diagrams.Gui;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.OrganizationChart;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class OrganizationDiagram extends OrganizationChart{


    private DiagramTree     m_DiagramTree   = null;

    
    public OrganizationDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        GROUPWIDTH          = 10;
        GROUPHEIGHT         = 6;
        WIDTH               = 4;
        HORSPACE            = 1;
        HEIGHT              = 4;
        VERSPACE            = 3;
    }
/*
    public OrganizationDiagram(Controller controller, Gui gui, XFrame xFrame, DiagramTree diagramTree) {
        this(controller, gui, xFrame);
        //m_DiagramTree = (ODiagramTree)diagramTree;
        m_DiagramTree = new ODiagramTree(this, diagramTree);
        System.out.println("init");
    }
*/
    @Override
    public DiagramTree getDiagramTree(){
        return m_DiagramTree;
    }
    
    @Override
    public String getDiagramType(){
        return "OrganizationDiagram";
    }

    @Override
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
            setPropsOfBaseShape(xBaseShape);

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
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } 
    }

    @Override
    public void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartShape", xStartShape);
            xProps.setPropertyValue("EndShape", xEndShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            xProps.setPropertyValue("LineWidth",new Integer(100));
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
    public void initDiagram() {
        //initial members: m_xDrawPage, m_DiagramID, m_xShapes
        super.initDiagram();
        if(m_DiagramTree == null)
            m_DiagramTree = new ODiagramTree(this);
        m_DiagramTree.setLists();
        m_DiagramTree.setTree();
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

                                    newTreeItem = new OTreeItem(m_DiagramTree, xRectangleShape, dadItem, level , (short)0);
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

                                    newTreeItem = new OTreeItem(m_DiagramTree, xRectangleShape, dadItem, level , (short)0);
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
            System.err.println(ex.getLocalizedMessage());
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
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
                    cutSelectedItem((OTreeItem) selectedItem);
                    getController().setSelectedShape((Object)xDadShape);

                }
            }
        }
    }

    public void cutSelectedItem(OTreeItem item){
        item.setDad(null);
        item.setFirstChild(null);
        item.setFirstSibling(null);
    }

}
