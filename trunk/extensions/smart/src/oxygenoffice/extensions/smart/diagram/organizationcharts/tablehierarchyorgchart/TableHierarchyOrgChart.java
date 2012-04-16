package oxygenoffice.extensions.smart.diagram.organizationcharts.tablehierarchyorgchart;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;


public class TableHierarchyOrgChart extends OrganizationChart {

    
    private OrganizationChartTree   m_DiagramTree   =  null;

    protected final static int GAP = 250;


    public TableHierarchyOrgChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth   = 10;
        m_iGroupHeight  = 6;
        m_iShapeWidth   = 2;
        m_iHorSpace     = 1;
        m_iShapeHeight  = 4;
        m_iVerSpace     = 3;
    }

    @Override
    public void initDiagramTree(OrganizationChartTree diagramTree){
        super.initDiagram();
        m_DiagramTree = new TableHierarchyOrgChartTree(this, diagramTree);
    }
 
    @Override
    public OrganizationChartTree getDiagramTree(){
        return m_DiagramTree;
    }

    @Override
    public String getDiagramTypeName(){
        return "TableHierarchyDiagram";
    }

    @Override
    public void createDiagram(int n){
        if(m_xDrawPage != null && m_xShapes != null && n > 0){
            setDrawArea();
            // base shape
            XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
            m_xShapes.add(xBaseShape);
            setControlShapeProps(xBaseShape);

            int horUnit, shapeWidth, verUnit, shapeHeight;
            horUnit = shapeWidth = verUnit =  shapeHeight = 0;
            shapeWidth = m_DrawAreaWidth;
            if(n > 1){
                horUnit = m_DrawAreaWidth / ( n - 1 );
                verUnit = m_DrawAreaHeight / 2;
                shapeHeight = m_DrawAreaHeight / 2 - GAP;
            }
            if(n == 1)
                shapeHeight = m_DrawAreaHeight;

            int xCoord = m_PageProps.BorderLeft + m_iHalfDiff;
            int yCoord = m_PageProps.BorderTop;

            XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
            m_xShapes.add(xStartShape);
            setTextOfShape(xStartShape, " ");
            setItemProperties(xStartShape, (short)0);
          
            yCoord += verUnit;
            shapeWidth = shapeWidth / 3 - GAP;
            XShape xRectShape = null;

            for( int i = 2; i <= n; i++ ){
                xRectShape = createShape("RectangleShape", i, xCoord + ((i-2) * horUnit), yCoord, shapeWidth, shapeHeight);
                m_xShapes.add(xRectShape);
                setTextOfShape(xRectShape, " ");
                setItemProperties(xRectShape, (short)1);

                XShape xConnectorShape = createShape("ConnectorShape", i);
                m_xShapes.add(xConnectorShape);
                setMoveProtectOfShape(xConnectorShape);
                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectShape, new Integer(0));
            }
            getController().setSelectedShape((Object)xStartShape);
        }
    }

    @Override
    public void initDiagram(){
        //initial members: m_xDrawPage, m_DiagramID, m_xShapes
        super.initDiagram();
        if(m_DiagramTree == null)
            m_DiagramTree = new TableHierarchyOrgChartTree(this);
        m_DiagramTree.setLists();
        m_DiagramTree.setTree();
    }

    @Override
    public void setConnectorShapeProps(XShape xConnectorShape, XShape xStartShape, Integer startIndex, XShape xEndShape, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartShape", xStartShape);
            xProps.setPropertyValue("EndShape", xEndShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            xProps.setPropertyValue("LineStyle", LineStyle.NONE);
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
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
                    OrganizationChartTreeItem selectedItem = m_DiagramTree.getTreeItem(xSelectedShape);
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
          
                            m_DiagramTree.addToRectangles(xRectangleShape);
                            OrganizationChartTreeItem newTreeItem = null;
                            OrganizationChartTreeItem dadItem = null;

                            short level = -1;
                            if(m_sNewItemHType == UNDERLING){
                                dadItem = selectedItem;
                                level = (short)(dadItem.getLevel() + 1);
                                newTreeItem = new TableHierarchyOrgChartTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0 , 0.0);
                                if(!dadItem.isFirstChild()){
                                    dadItem.setFirstChild(newTreeItem);
                                }else{
                                    XShape xPreviousChild = m_DiagramTree.getLastChildShape(xSelectedShape);
                                    if(xPreviousChild != null){
                                        OrganizationChartTreeItem previousItem = m_DiagramTree.getTreeItem(xPreviousChild);
                                        if(previousItem != null)
                                            previousItem.setFirstSibling(newTreeItem);
                                    }
                                }
                            }else if(m_sNewItemHType == ASSOCIATE){
                                level = selectedItem.getLevel();
                                dadItem = selectedItem.getDad();
                                newTreeItem = new TableHierarchyOrgChartTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0, 0.0);
                                if(selectedItem.isFirstSibling())
                                    newTreeItem.setFirstSibling(selectedItem.getFirstSibling());
                                selectedItem.setFirstSibling(newTreeItem);
                            }

                            setTextOfShape(xRectangleShape, " ");
                            setItemProperties(xRectangleShape, level);
                            
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
                }
            }
        }
    }

    @Override
    public void setConnectorShapeProps(XShape xConnectorShape, Integer startIndex, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            //xProps.setPropertyValue("LineWidth",new Integer(100));
            xProps.setPropertyValue("LineStyle", LineStyle.NONE);
            //xProps.setPropertyValue("EdgeKind", ConnectorType.STANDARD);
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

}
