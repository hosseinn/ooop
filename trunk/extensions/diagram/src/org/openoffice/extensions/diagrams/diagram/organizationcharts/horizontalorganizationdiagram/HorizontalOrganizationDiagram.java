package org.openoffice.extensions.diagrams.diagram.organizationcharts.horizontalorganizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.ConnectorType;
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


public class HorizontalOrganizationDiagram extends OrganizationChart{


    private DiagramTree        m_DiagramTree   =  null;


    public HorizontalOrganizationDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        GROUPWIDTH      = 4;
        GROUPHEIGHT     = 3;
        WIDTH           = 2;
        HORSPACE        = 1;
        HEIGHT          = 4;
        VERSPACE        = 3;
    }

    @Override
    public DiagramTree getDiagramTree(){
        return m_DiagramTree;
    }

    @Override
    public String getDiagramType(){
        return "HorizontalOrganizationDiagram";
    }

    @Override
    public void createDiagram(int n){
        try{
            if(m_xDrawPage != null && m_xShapes != null && n > 0){
         
                int orignGSWidth = m_GroupSizeWidth;
    
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
                XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderBottom, m_GroupSizeWidth, m_GroupSizeHeight);
                m_xShapes.add(xBaseShape);
                setPropsOfBaseShape(xBaseShape);
            

                int horUnit = m_GroupSizeWidth / ( 2 * WIDTH + HORSPACE );
                int horSpace = horUnit * HORSPACE;
                int shapeWidth = horUnit * WIDTH;

                int verUnit = m_GroupSizeHeight / ( 3 * HEIGHT + 2 * VERSPACE);
                int verSpace = verUnit * VERSPACE;
                int shapeHeight = verUnit * HEIGHT;

                int xCoord = m_PageProps.BorderLeft + halfDiff;
                int yCoord = m_PageProps.BorderTop +shapeHeight +verSpace;

                XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
                m_xShapes.add(xStartShape);
                setMoveProtectOfShape(xStartShape);
                setTextFitToSize(xStartShape);
                setShapeProperties(xStartShape, "RectangleShape");
                setColorOfShape(xStartShape, COLOR);

                xCoord += (shapeWidth + horSpace);
                yCoord = m_PageProps.BorderTop;
                XShape xRectShape = null;
    

                for( int i = 2; i <= n; i++ ){
                    xRectShape = createShape("RectangleShape", i, xCoord, yCoord + m_GroupSizeHeight - shapeHeight - (shapeHeight + verSpace) * (i-2), shapeWidth, shapeHeight);
                    m_xShapes.add(xRectShape);
                    setMoveProtectOfShape(xRectShape);
                    setTextFitToSize(xRectShape);
                    setShapeProperties(xRectShape, "RectangleShape");
                    setColorOfShape(xRectShape, COLOR);

                    XShape xConnectorShape = createShape("ConnectorShape", i);
                    m_xShapes.add(xConnectorShape);
                    setMoveProtectOfShape(xConnectorShape);
                    setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(1), xRectShape, new Integer(3));

                }
                getController().setSelectedShape((Object)xStartShape);
            }
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void initDiagram() {
        //initial members: m_xDrawPage, m_DiagramID, m_xShapes
        super.initDiagram();
        if(m_DiagramTree == null)
            m_DiagramTree = new HDiagramTree(this);
        m_DiagramTree.setLists();
        m_DiagramTree.setTree();
    }

    @Override
    public void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProp.setPropertyValue("StartShape", xStartShape);
            xProp.setPropertyValue("EndShape", xEndShape);
            xProp.setPropertyValue("StartGluePointIndex", startIndex);
            xProp.setPropertyValue("EndGluePointIndex", endIndex);
            xProp.setPropertyValue("EdgeKind", ConnectorType.LINE);
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

                    if(selectedShapeName.contains("RectangleShape") && !selectedShapeName.contains("RectangleShape0")){
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
                                m_xShapes.add(xRectangleShape);
                                setMoveProtectOfShape(xRectangleShape);
                                setTextFitToSize(xRectangleShape);
                                setShapeProperties(xRectangleShape, "RectangleShape");
                                int color = getGui().getImageColorOfControlDialog();
                                if(color < 0)
                                    color = COLOR;
                                setColorOfShape(xRectangleShape, color);

                                m_DiagramTree.addToRectangles(xRectangleShape);

                                TreeItem newTreeItem = null;
                                TreeItem dadItem = null;

                                if(m_sNewItemHType == UNDERLING){
                                    dadItem = selectedItem;
                                    newTreeItem = new HTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0 , 0.0);
                                    if(!dadItem.isFirstChild()){
                                        dadItem.setFirstChild(newTreeItem);
                                    }else{
                                        XShape xPreviousChild = m_DiagramTree.getLastChildShape(xSelectedShape);
                                        if(xPreviousChild != null){
                                            TreeItem previousItem = m_DiagramTree.getTreeItem(xPreviousChild);
                                            if(previousItem != null)
                                                previousItem.setFirstSibling(newTreeItem);
                                        }
                                    }

                                }else if(m_sNewItemHType == ASSOCIATE){
                                    dadItem = selectedItem.getDad();
                                    newTreeItem = new HTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0, 0.0);
                                    if(selectedItem.isFirstSibling())
                                        newTreeItem.setFirstSibling(selectedItem.getFirstSibling());
                                    selectedItem.setFirstSibling(newTreeItem);
                                }

                                if(iTopShapeID > 1){
                                    // set connector shape
                                    XShape xConnectorShape = createShape("ConnectorShape", iTopShapeID);
                                    m_xShapes.add(xConnectorShape);
                                    setMoveProtectOfShape(xConnectorShape);
                                    m_DiagramTree.addToConnectors(xConnectorShape);

                                    XShape xStartShape = null;
                                    if(m_sNewItemHType == UNDERLING)
                                        xStartShape = getController().getSelectedShape();
                                    else if(m_sNewItemHType == ASSOCIATE)
                                        xStartShape = selectedItem.getDad().getRectangleShape();
                                    setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(1), xRectangleShape, new Integer(3));
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
                    }

                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        }  catch (WrappedTargetException ex) {
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
                    setNullSelectedItem(selectedItem);
                    getController().setSelectedShape((Object)xDadShape);
                }
            }
        }
    }

}
