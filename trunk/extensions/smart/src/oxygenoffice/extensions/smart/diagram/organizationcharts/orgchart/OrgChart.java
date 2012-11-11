package oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart;

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


public class OrgChart extends OrganizationChart{


    private OrganizationChartTree     m_DiagramTree   = null;

    
    public OrgChart(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        m_iGroupWidth       = 10;
        m_iGroupHeight      = 6;
        m_iShapeWidth       = 4;
        m_iHorSpace         = 1;
        m_iShapeHeight      = 4;
        m_iVerSpace         = 3;
    }

    @Override
    public void initDiagramTree(OrganizationChartTree diagramTree){
        super.initDiagram();
        m_DiagramTree = new OrgChartTree(this, diagramTree);
    }

    @Override
    public OrganizationChartTree getDiagramTree(){
        return m_DiagramTree;
    }

    @Override
    public String getDiagramTypeName(){
        return "OrganizationDiagram";
    }

    @Override
    public void createDiagram(DataOfDiagram datas){
        short lastHorLevel = OrgChartTree.LASTHORLEVEL;
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
                
                OrgChartTree.LASTHORLEVEL = (short)10000;
                setHorLevelOfControlShape(xBaseShape, (short)10000);

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

                m_DiagramTree = new OrgChartTree(this, xBaseShape, xStartShape);
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
                        int lev = dadItem.getLevel() + 1 - datas.get(i).getLevel();
                        for(int j = 0; j < lev; j++)
                            dadItem = dadItem.getDad();
                    }

                    XShape xConnectorShape = createShape("ConnectorShape", i + (2 - iRoot));
                    m_xShapes.add(xConnectorShape);
                    setMoveProtectOfShape(xConnectorShape);
                    Integer endShapeConnPos = new Integer(0);
                    if(dadItem.getLevel() + 1 > lastHorLevel)
                        endShapeConnPos = new Integer(3);
                    setConnectorShapeProps(xConnectorShape, dadItem.getRectangleShape(), new Integer(2), xShape, endShapeConnPos);
                    m_DiagramTree.addToConnectors(xConnectorShape);

                    newTreeItem = new OrgChartTreeItem(m_DiagramTree, xShape, dadItem, (short)0, 0.0);
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
                    //refreshDiagram();
                }else{
                    iColor++;
                    iColor %= 5;
                    setColorProp(aLOCOLORS2[iColor][1]);
                    getController().setSelectedShape((Object)m_DiagramTree.getRootItem().getRectangleShape());

                }
                OrgChartTree.LASTHORLEVEL = lastHorLevel;
                setHorLevelOfControlShape(xBaseShape, lastHorLevel);
                refreshDiagram();
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
            
            OrgChartTree.LASTHORLEVEL = 2;
            setHorLevelOfControlShape(xBaseShape, (short)2);
            //XText xText = UnoRuntime.queryInterface(XText.class, xBaseShape);
            //xText.setString("ColorMode:0:Style:0:LastHorLevel:2");

            int horUnit, horSpace, shapeWidth, verUnit, verSpace, shapeHeight;
            horUnit = horSpace = shapeWidth = verUnit = verSpace =  shapeHeight = 0;
            if(n > 1){
                horUnit = m_DrawAreaWidth / ( (n-1) * m_iShapeWidth + (n-2) * m_iHorSpace );
                horSpace = horUnit * m_iHorSpace;
                shapeWidth = horUnit * m_iShapeWidth;
                verUnit = m_DrawAreaHeight / ( 2 * m_iShapeHeight + m_iVerSpace);
                verSpace = verUnit * m_iVerSpace;
                shapeHeight = verUnit * m_iShapeHeight;
            }
            if(n == 1){
                shapeWidth = m_DrawAreaWidth;
                shapeHeight = m_DrawAreaHeight;
            }

            int xCoord = m_PageProps.BorderLeft + m_iHalfDiff + m_DrawAreaWidth / 2 - shapeWidth / 2;
            int yCoord = m_PageProps.BorderTop;

            XShape xStartShape = createShape("RectangleShape", 1, xCoord, yCoord, shapeWidth, shapeHeight);
            m_xShapes.add(xStartShape);
            setTextOfShape(xStartShape, " ");
            setMoveProtectOfShape(xStartShape);
            setColorProp(aORGCHARTCOLORS[0]);
            setShapeProperties(xStartShape, "RectangleShape", true);

            xCoord = m_PageProps.BorderLeft + m_iHalfDiff;
            yCoord = m_PageProps.BorderTop + shapeHeight + verSpace;
            XShape xRectShape = null;
            XShape xSelectedShape = null;

            int i;
            for( i = 2; i <= n; i++ ){
                xRectShape = createShape("RectangleShape", i, xCoord + (shapeWidth + horSpace) * (i-2), yCoord, shapeWidth, shapeHeight);
                m_xShapes.add(xRectShape);
                setTextOfShape(xRectShape, " ");
                setMoveProtectOfShape(xRectShape);
                setColorProp(aORGCHARTCOLORS[(i - 1) % 8]);
                setShapeProperties(xRectShape, "RectangleShape", true);

                XShape xConnectorShape = createShape("ConnectorShape", i);
                m_xShapes.add(xConnectorShape);
                setMoveProtectOfShape(xConnectorShape);
                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectShape, new Integer(0));
                if(i == 2 && xRectShape != null)
                    xSelectedShape = xRectShape;
            }
            int id = 1;
            if(n == 1 && xStartShape != null){
                getController().setSelectedShape((Object)xStartShape);
            } else if(xSelectedShape != null){
                getController().setSelectedShape((Object)xSelectedShape);
                id = getController().getShapeID(getShapeName(xSelectedShape));
            }
            setColorProp(aORGCHARTCOLORS[(id - 1) % 8]);
        }
    }

    @Override
    public void initDiagram() {
        //initial members: m_xDrawPage, m_DiagramID, m_xShapes
        super.initDiagram();
        if(m_DiagramTree == null)
            m_DiagramTree = new OrgChartTree(this);
        m_DiagramTree.setLists();
        if(OrgChartTree.LASTHORLEVEL == -1)
            OrgChartTree.LASTHORLEVEL = getHorLevelOfControlShape(m_DiagramTree.getControlShape());
        else
            setHorLevelOfControlShape(m_DiagramTree.getControlShape(), OrgChartTree.LASTHORLEVEL);
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

                            if(m_sNewItemHType == UNDERLING){
                                dadItem = selectedItem;
                                newTreeItem = new OrgChartTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0 , 0.0);
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
                                dadItem = selectedItem.getDad();
                                newTreeItem = new OrgChartTreeItem(m_DiagramTree, xRectangleShape, dadItem, (short)0, 0.0);
                                if(selectedItem.isFirstSibling())
                                   newTreeItem.setFirstSibling(selectedItem.getFirstSibling());
                                selectedItem.setFirstSibling(newTreeItem);
                            }

                            setTextOfShape(xRectangleShape, " ");
                            setMoveProtectOfShape(xRectangleShape);
                            //setFontPropertiesOfShape(xRectangleShape);
                            //setColorSettingsOfShape(xRectangleShape);
                            setShapeProperties(xRectangleShape, "RectangleShape", true);
                            
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
                                    if(selectedItem.getLevel() + 1 > OrgChartTree.LASTHORLEVEL)
                                       endShapeConnPos = new Integer(3);
                                }else if (m_sNewItemHType == ASSOCIATE) {
                                    xStartShape = selectedItem.getDad().getRectangleShape();
                                    if(selectedItem.getLevel() > OrgChartTree.LASTHORLEVEL)
                                       endShapeConnPos = new Integer(3);
                                }
                                setConnectorShapeProps(xConnectorShape, xStartShape, new Integer(2), xRectangleShape, endShapeConnPos);
                                if(isHiddenRootElementProp())
                                    if(getDiagramTree().getRootItem().getRectangleShape().equals(xStartShape))
                                        getDiagramTree().getRootItem().hideElement();
                            }
                        }
                    }
                }
            }
        }   
    }
    
}
