package org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.ConnectorType;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
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

    @Override
    public void initDiagramTree(DiagramTree diagramTree){
        super.initDiagram();
        m_DiagramTree = new ODiagramTree(this, diagramTree);
    }

    @Override
    public DiagramTree getDiagramTree(){
        return m_DiagramTree;
    }
    
    @Override
    public String getDiagramTypeName(){
        return "OrganizationDiagram";
    }

    @Override
    public void createDiagram(int n){
        if(m_xDrawPage != null && m_xShapes != null && n > 0){
            setDrawArea();
            // base shape
            XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderBottom, m_DrawAreaWidth, m_DrawAreaHeight);
            m_xShapes.add(xBaseShape);
            setPropsOfBaseShape(xBaseShape);

            int horUnit, horSpace, shapeWidth, verUnit, verSpace, shapeHeight;
            horUnit = horSpace = shapeWidth = verUnit = verSpace =  shapeHeight = 0;
            if(n > 1){
                horUnit = m_DrawAreaWidth / ( (n-1) * WIDTH + (n-2) * HORSPACE );
                horSpace = horUnit * HORSPACE;
                shapeWidth = horUnit * WIDTH;
                verUnit = m_DrawAreaHeight / ( 2 * HEIGHT + VERSPACE);
                verSpace = verUnit * VERSPACE;
                shapeHeight = verUnit * HEIGHT;
            }
            if(n == 1){
                shapeWidth = m_DrawAreaWidth;
                shapeHeight = m_DrawAreaHeight;
            }

            int xCoord = m_PageProps.BorderLeft + m_iHalfDiff + m_DrawAreaWidth / 2 - shapeWidth / 2;
            int yCoord = m_PageProps.BorderTop;

            XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
            m_xShapes.add(xStartShape);
            setMoveProtectOfShape(xStartShape);
            setTextFitToSize(xStartShape);
            setShapeProperties(xStartShape, "RectangleShape");
            int color = -1;
            if(getGui() != null && getGui().getControlDialogWindow() != null)
                color = getGui().getImageColorOfControlDialog();
            if(color < 0)
                color = COLOR;
            setColorOfShape(xStartShape, color);

            xCoord = m_PageProps.BorderLeft + m_iHalfDiff;
            yCoord = m_PageProps.BorderTop + shapeHeight + verSpace;
            XShape xRectShape = null;

            for( int i = 2; i <= n; i++ ){
                xRectShape = createShape("RectangleShape", i, xCoord + (shapeWidth + horSpace) * (i-2), yCoord, shapeWidth, shapeHeight);
                m_xShapes.add(xRectShape);
                setMoveProtectOfShape(xRectShape);
                setTextFitToSize(xRectShape);
                setShapeProperties(xRectShape, "RectangleShape");
                setColorOfShape(xRectShape, color);

                XShape xConnectorShape = createShape("ConnectorShape", i);
                m_xShapes.add(xConnectorShape);
                setMoveProtectOfShape(xConnectorShape);
                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectShape, new Integer(0));
            }
            getController().setSelectedShape((Object)xStartShape);
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
    public void setConnectorShapeProps(XShape xConnectorShape, Integer startIndex, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            xProps.setPropertyValue("LineWidth",new Integer(100));
            xProps.setPropertyValue("LineStyle", LineStyle.SOLID);
            xProps.setPropertyValue("EdgeKind", ConnectorType.STANDARD);
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
        if(m_DiagramTree != null){
            XShape xSelectedShape = getController().getSelectedShape();
            if(xSelectedShape != null){
                XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSelectedShape );
                String selectedShapeName = xNamed.getName();
                int iTopShapeID = -1;

                if(selectedShapeName.contains("RectangleShape") && !selectedShapeName.contains("RectangleShape0")){
                    TreeItem selectedItem = m_DiagramTree.getTreeItem(xSelectedShape);
                    // can't be associate of root item
                    if(selectedItem.getDad() == null && m_sNewItemHType == 1 ) {
                        String title = getGui().getDialogPropertyValue("Strings", "ItemAddError.Title");
                        String message = getGui().getDialogPropertyValue("Strings", "ItemAddError.Message");
                        getGui().showMessageBox(title, message);
                    }else{
                        iTopShapeID = getTopShapeID();
                        if( iTopShapeID <= 0 ){
                            clearEmptyDiagramAndReCreate();
                        }else{
                            iTopShapeID++;
                            XShape xRectangleShape = createShape("RectangleShape", iTopShapeID);
                            m_xShapes.add(xRectangleShape);
                            setMoveProtectOfShape(xRectangleShape);
                            setTextFitToSize(xRectangleShape);
                            setShapeProperties(xRectangleShape, "RectangleShape");
                            int color = -1;
                            if(getGui() != null && getGui().getControlDialogWindow() != null)
                                color = getGui().getImageColorOfControlDialog();
                            if(color < 0)
                                color = COLOR;
                            setColorOfShape(xRectangleShape, color);
                            m_DiagramTree.addToRectangles(xRectangleShape);

                            TreeItem newTreeItem = null;
                            TreeItem dadItem = null;

                            if(m_sNewItemHType == UNDERLING){
                                dadItem = selectedItem;
                                newTreeItem = new OTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0 , 0.0);
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
                                newTreeItem = new OTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0, 0.0);
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
                                Integer endShapeConnPos = new Integer(0);
                                if(m_sNewItemHType == UNDERLING){
                                    xStartShape = selectedItem.getRectangleShape();
                                    if(selectedItem.getLevel() + 1 > ODiagramTree.LASTHORLEVEL)
                                       endShapeConnPos = new Integer(3);
                                }else if (m_sNewItemHType == ASSOCIATE) {
                                    xStartShape = selectedItem.getDad().getRectangleShape();
                                    if(selectedItem.getLevel() > ODiagramTree.LASTHORLEVEL)
                                       endShapeConnPos = new Integer(3);
                                }
                                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectangleShape, endShapeConnPos);
                            }
                        }
                    }
                }
            }
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
