package org.openoffice.extensions.diagrams.diagram.organizationcharts.tablehierarchydiagram;

import com.sun.star.awt.Point;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


class THDiagramTree extends DiagramTree {


    public THDiagramTree(TableHierarchyDiagram thOrganigram) {
        super(thOrganigram);
    }

    @Override
    public void initTreeItems(){
        THTreeItem.initStaticMembers();
        m_RootItem = new THTreeItem( this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public void refresh(){
        THTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        ((THTreeItem)m_RootItem).setWidthUnit(1.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
        m_RootItem.display();
    }

    @Override
    public XShape getFirstChildShape(XShape xDadShape){
        int     xPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( xPos == -1 || xChildeShape.getPosition().X < xPos){
                    xPos = xChildeShape.getPosition().X;
                    xFirstChildShape = xChildeShape;
                }
            }
        }
        return xFirstChildShape;
    }

    @Override
    public XShape getFirstSiblingShape(XShape xBaseShape, TreeItem dad){
        if(dad == null)
            return null;
        if(dad.getRectangleShape() == null)
            return null;
        XShape xDadShape            = dad.getRectangleShape();
        XShape xSiblingShape        = null;
        XShape xFirstSiblingShape   = null;
        Point  baseShapePos         = xBaseShape.getPosition();
        int    xPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if( xSiblingShape.getPosition().X > baseShapePos.X){
                    if( xPos == -1 || xSiblingShape.getPosition().X < xPos){
                        xPos = xSiblingShape.getPosition().X;
                        xFirstSiblingShape = xSiblingShape;
                    }
                }

            }
        }
        return xFirstSiblingShape;
    }

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     xPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( xPos == -1 || xChildeShape.getPosition().X > xPos){
                    xPos = xChildeShape.getPosition().X;
                    xLastChildShape = xChildeShape;
                }
            }
        }
        return xLastChildShape;
    }

}
