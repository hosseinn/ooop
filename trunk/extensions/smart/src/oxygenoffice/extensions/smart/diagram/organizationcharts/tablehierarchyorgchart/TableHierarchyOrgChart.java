package oxygenoffice.extensions.smart.diagram.organizationcharts.tablehierarchyorgchart;

import com.sun.star.container.XNamed;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.DataOfDiagram;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class TableHierarchyOrgChart extends OrganizationChart {

    
    private OrganizationChartTree   m_DiagramTree   =  null;

    protected final static int GAP = 150;


    public TableHierarchyOrgChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        setShownConnectorsProp(false);
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
    public void createDiagram(DataOfDiagram datas){
        if(!datas.isEmpty()){
            super.createDiagram(datas);
            boolean isRootItem = datas.isOneFirstLevelData();
            if(!isRootItem)
                datas.increaseLevels();
            if(m_xDrawPage != null && m_xShapes != null){
                setDrawArea();
                // base shape
                XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
                m_xShapes.add(xBaseShape);
                setControlShapeProps(xBaseShape);
                setColorModeAndStyeOfControlShape(xBaseShape);

                XShape xStartShape = createShape("RectangleShape", 1, m_PageProps.BorderLeft, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
                m_xShapes.add(xStartShape);
                if(isRootItem)
                    setTextOfShape(xStartShape, datas.get(0).getValue());
                else
                    setTextOfShape(xStartShape, " ");
                setMoveProtectOfShape(xStartShape);
                setColorProp(aLOORANGES[2]);
                setShapeProperties(xStartShape, "RectangleShape", true);
                if(xStartShape != null)
                    getController().setSelectedShape((Object)xStartShape);
                initDiagram();

                m_DiagramTree = new TableHierarchyOrgChartTree(this, xBaseShape, xStartShape);
                OrganizationChartTreeItem dadItem = m_DiagramTree.getRootItem();
                OrganizationChartTreeItem newTreeItem = null;
                OrganizationChartTreeItem lastTreeItem = dadItem;
                int size = datas.size();
                int iRoot = 0;
                if(isRootItem)
                    iRoot = 1;
                int iColor = 0;
                for( int i = iRoot; i < size; i++){
                    XShape xShape = createShape("RectangleShape", i + (2 - iRoot));
                    m_xShapes.add(xShape);
                    setTextOfShape(xShape, datas.get(i).getValue());
                    setMoveProtectOfShape(xShape);
                    if(i > iRoot)
                        if(datas.get(i).getLevel() == 1)
                            iColor++;
                    iColor %= 5;
                    int iColorLevel = datas.get(i).getLevel();
                    if(iColorLevel > 4)
                        iColorLevel = 4;
                    setColorProp(aLOCOLORS2[iColor][iColorLevel]);
                    setShapeProperties(xShape, "RectangleShape", true);
                    m_DiagramTree.addToRectangles(xShape);

                    if(lastTreeItem.getLevel() == datas.get(i).getLevel()){
                    }else if(lastTreeItem.getLevel() < datas.get(i).getLevel()){
                        dadItem = lastTreeItem;
                    }else{
                        int lev = dadItem.getLevel() + 1 - datas.get(i).getLevel();;
                        for(int j = 0; j < lev; j++)
                            dadItem = dadItem.getDad();
                    }

                    XShape xConnectorShape = createShape("ConnectorShape", i + (2 - iRoot));
                    m_xShapes.add(xConnectorShape);
                    setMoveProtectOfShape(xConnectorShape);
                    setConnectorShapeProps(xConnectorShape, dadItem.getRectangleShape(), new Integer(2), xShape, new Integer(0));
                    m_DiagramTree.addToConnectors(xConnectorShape);

                    newTreeItem = new TableHierarchyOrgChartTreeItem(m_DiagramTree, xShape, dadItem, (short)0, 0.0);
                    if(lastTreeItem.getLevel() == datas.get(i).getLevel()){
                        lastTreeItem.setFirstSibling(newTreeItem);
                    }else if(lastTreeItem.getLevel() < datas.get(i).getLevel()){
                        if(!dadItem.isFirstChild()){
                            dadItem.setFirstChild(newTreeItem);
                        }
                    }else{
                        dadItem.getLastChild().setFirstSibling(newTreeItem);
                    }
                    lastTreeItem = newTreeItem;
                    refreshDiagram();
                }
                if(!isRootItem){
                    getController().setSelectedShape((Object)m_DiagramTree.getRootItem().getLastChild().getRectangleShape());
                    setHiddenRootElementProp(true);
                    getDiagramTree().getRootItem().hideElement();
                    refreshDiagram();
                }else{
                    iColor++;
                    iColor %= 5;
                    setColorProp(aLOCOLORS2[iColor][1]);
                    getController().setSelectedShape((Object)m_DiagramTree.getRootItem().getRectangleShape());

                }
            }
        }
    }

    @Override
    public void createDiagram(int n){
        if(m_xDrawPage != null && m_xShapes != null && n > 0){
            setDrawArea();
            // base shape
            XShape xBaseShape = createShape("RectangleShape", 0, m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop, m_DrawAreaWidth, m_DrawAreaHeight);
            m_xShapes.add(xBaseShape);
            setControlShapeProps(xBaseShape);
            setColorModeAndStyeOfControlShape(xBaseShape);

            int horUnit, shapeWidth, verUnit, shapeHeight;
            horUnit = shapeWidth = verUnit =  shapeHeight = 0;
            shapeWidth = m_DrawAreaWidth - GAP - getShapesLineWidhtProp();
            if(n > 1){
                horUnit = m_DrawAreaWidth / ( n - 1 );
                verUnit = m_DrawAreaHeight / 2;
                shapeHeight = m_DrawAreaHeight / 2 - GAP - getShapesLineWidhtProp();
                if(isShadowProp())
                    shapeHeight -= OrganizationChart.SHADOW_DIST1;
            }
            if(n == 1)
                shapeHeight = m_DrawAreaHeight;

            int xCoord = m_PageProps.BorderLeft + m_iHalfDiff;
            int yCoord = m_PageProps.BorderTop;

            XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
            m_xShapes.add(xStartShape);
            setTextOfShape(xStartShape, " ");
            setMoveProtectOfShape(xStartShape);
            setColorProp(aLOGREENS[0]);
            setShapeProperties(xStartShape, "RectangleShape", true);
          
            yCoord += verUnit;
            shapeWidth = shapeWidth / 3 - GAP - getShapesLineWidhtProp();
            if(isShadowProp())
                    shapeWidth -= OrganizationChart.SHADOW_DIST1;
            XShape xRectShape = null;
            XShape xSelectedShape = null;

            for( int i = 2; i <= n; i++ ){
                xRectShape = createShape("RectangleShape", i, xCoord + ((i-2) * horUnit), yCoord, shapeWidth, shapeHeight);
                m_xShapes.add(xRectShape);
                setTextOfShape(xRectShape, " ");
                setMoveProtectOfShape(xRectShape);
                setColorProp(aLOGREENS[1]);
                setShapeProperties(xRectShape, "RectangleShape", true);

                XShape xConnectorShape = createShape("ConnectorShape", i);
                m_xShapes.add(xConnectorShape);
                setMoveProtectOfShape(xConnectorShape);
                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectShape, new Integer(0));
                if(i == 2 && xRectShape != null)
                    xSelectedShape = xRectShape;
            }
            if(n == 1 && xStartShape != null){
                getController().setSelectedShape((Object)xStartShape);
            } else if(xSelectedShape != null){
                getController().setSelectedShape((Object)xSelectedShape);
            }
            setColorProp(aLOGREENS[2]);
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
/*
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
*/
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

                            //short level = -1;
                            if(m_sNewItemHType == UNDERLING){
                                dadItem = selectedItem;
                                //level = (short)(dadItem.getLevel() + 1);
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
                                //level = selectedItem.getLevel();
                                dadItem = selectedItem.getDad();
                                newTreeItem = new TableHierarchyOrgChartTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0, 0.0);
                                if(selectedItem.isFirstSibling())
                                    newTreeItem.setFirstSibling(selectedItem.getFirstSibling());
                                selectedItem.setFirstSibling(newTreeItem);
                            }

                            setTextOfShape(xRectangleShape, " ");
                            setMoveProtectOfShape(xRectangleShape);
//                            setFontPropertiesOfShape(xRectangleShape);
//                            setColorSettingsOfShape(xRectangleShape);
                            setShapeProperties(xRectangleShape, "RectangleShape", true);
                            
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
         //                       if(isHiddenRootElementProps())
         //                           if(getDiagramTree().getRootItem().getRectangleShape().equals(xStartShape))
         //                               getDiagramTree().getRootItem().hideElement();
                            }
                        }
                    }
                }
            }
        }
    }
/*
    @Override
    public void setConnectorShapeProps(XShape xConnectorShape, Integer startIndex, Integer endIndex){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnectorShape);
            xProps.setPropertyValue("StartGluePointIndex", startIndex);
            xProps.setPropertyValue("EndGluePointIndex", endIndex);
            xProps.setPropertyValue("LineStyle", LineStyle.NONE);
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
*/
}
