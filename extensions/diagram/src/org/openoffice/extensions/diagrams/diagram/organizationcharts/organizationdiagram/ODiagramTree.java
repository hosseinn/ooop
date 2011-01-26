package org.openoffice.extensions.diagrams.diagram.organizationcharts.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class ODiagramTree extends DiagramTree{


    public ODiagramTree(OrganizationDiagram organigram){
        super(organigram);
    }
/*
    ODiagramTree(OrganizationDiagram organigram, DiagramTree diagramTree) {
        this(organigram);

        rectangleList       = diagramTree.rectangleList;
        connectorList       = diagramTree.connectorList;
        m_xShapes           = diagramTree.m_xShapes;
        m_xControlShape     = diagramTree.m_xControlShape;
        m_xRootShape        = diagramTree.m_xRootShape;
        m_RootItem          = diagramTree.m_RootItem;
        m_SelectedItem      = diagramTree.m_SelectedItem;
    }
*/
    @Override
    public void initTreeItems(){
        OTreeItem.initStaticMembers();
        m_RootItem = new OTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public XShape getFirstChildShape(XShape xDadShape){
        // the struct of diagram change below second level that's why we need level of shape
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xPos == -1 || xChildeShape.getPosition().X < xPos){
                        xPos = xChildeShape.getPosition().X;
                        xFirstChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y < yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xFirstChildShape = xChildeShape;
                    }
                }
            }
        }
        return xFirstChildShape;
    }

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xPos == -1 || xChildeShape.getPosition().X > xPos){
                        xPos = xChildeShape.getPosition().X;
                        xLastChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y > yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xLastChildShape = xChildeShape;
                    }

                }
            }
        }
        return xLastChildShape;
    }

    @Override
    public XShape getFirstSiblingShape(XShape xBaseShape, TreeItem dad){
        if(dad == null)
            return null;
        if(dad.getRectangleShape() == null)
            return null;
        int    level                = dad.getLevel() + 1;
        XShape xDadShape            = dad.getRectangleShape();
        XShape xSiblingShape        = null;
        XShape xFirstSiblingShape   = null;
        Point  baseShapePos         = xBaseShape.getPosition();
        int    xPos                 = -1;
        int    yPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xSiblingShape.getPosition().X > baseShapePos.X){
                        if( xPos == -1 || xSiblingShape.getPosition().X < xPos){
                            xPos = xSiblingShape.getPosition().X;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }else{
                    if( xSiblingShape.getPosition().Y > baseShapePos.Y){
                        if( yPos == -1 || xSiblingShape.getPosition().Y < yPos){
                            yPos = xSiblingShape.getPosition().Y;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }
            }
        }
        return xFirstSiblingShape;
    }

    @Override
    public void refresh(){
        OTreeItem.initStaticMembers();
        TreeItem._maxLevel = 0;
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
        m_RootItem.display();
    }

}
