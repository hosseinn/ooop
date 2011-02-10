package org.openoffice.extensions.diagrams.diagram.organizationcharts.tablehierarchydiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


class THTreeItem extends TreeItem{


    private static double[]     _maxPositions;
    private static double       _maxPos             = -1.0;

    private double              m_WidthUnit         = 1.0;
    private double              m_LastSiblingUnit   = 1.0;

    private static int          _shapeWidth         = 0;
    private static int          _shapeHeight        = 0;
    private static int          _groupPosX          = 0;
    private static int          _groupPosY          = 0;


    THTreeItem(DiagramTree diagramTree, TreeItem dad, TreeItem item) {
        super(diagramTree, dad, item);
    }
    
    public THTreeItem(DiagramTree diagramTree, XShape xShape, TreeItem dad, short level, double num){
        super(diagramTree, xShape, dad);
        setLevel(level);
        setPos(num);
        m_WidthUnit         = 1.0;
        m_LastSiblingUnit   = 1.0;
    }

    public static void initStaticMembers(){
        _maxLevel = -1;
        _maxPos = -1.0;
        _maxPositions = new double[100];
        for(int i=0; i<_maxPositions.length; i++)
            _maxPositions[i] = -1.0;
    }

    @Override
    public void convertTreeItems(TreeItem treeItem){
        if(treeItem.isFirstChild()){
            m_FirstChild = new THTreeItem(getDiagramTree(), this, treeItem.getFirstChild());
            m_FirstChild.convertTreeItems(treeItem.getFirstChild());
        }
        if(treeItem.isFirstSibling()){
            m_FirstSibling = new THTreeItem(getDiagramTree(), getDad(), treeItem.getFirstSibling());
            m_FirstSibling.convertTreeItems(treeItem.getFirstSibling());
        }
    }

    public double getWidthUnit(){
        return m_WidthUnit;
    }

    public void setWidthUnit(double unit){
        m_WidthUnit = unit;
    }

    public double getLastSiblingUnit(){
        return m_LastSiblingUnit;
    }

    public void setLastSiblingUnit(double unit){
        m_LastSiblingUnit = unit;
    }

    @Override
    public final void setPos(double pos){
        m_Pos = pos;
        if(m_Pos > _maxPositions[m_Level])
            _maxPositions[m_Level] = m_Pos;
        if(m_Pos > _maxPos)
            _maxPos = m_Pos;
        //XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
        //xText.setString(getWidthUnit() + " : " + (int)getLastSiblingUnit());
    }

    @Override
    public void initTreeItems(){
        XShape xFirstChildShape = getDiagramTree().getFirstChildShape(m_xRectangleShape);
        if(xFirstChildShape != null){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = _maxPositions[firstChildLevel] + 1.0;
            m_FirstChild = new THTreeItem(getDiagramTree(), xFirstChildShape, this, firstChildLevel , firstChildPos);
            m_FirstChild.initTreeItems();
        }

        XShape xFirstSiblingShape = getDiagramTree().getFirstSiblingShape(m_xRectangleShape, m_Dad);
        if(xFirstSiblingShape != null){
            short firstSiblingLevel = m_Level;
            double firstSiblingNum = m_Pos + 1.0;
            m_FirstSibling = new THTreeItem(getDiagramTree(), xFirstSiblingShape, m_Dad, firstSiblingLevel , firstSiblingNum);
            m_FirstSibling.initTreeItems();
            setLastSiblingUnit(((THTreeItem)m_FirstSibling).getLastSiblingUnit());
        }else{
            setLastSiblingUnit(getWidthUnit());
        }

        if(m_Dad != null && m_Dad.getFirstChild().equals(this)){
            if(isFirstSibling())
                ((THTreeItem)m_Dad).setWidthUnit(_maxPositions[m_Level] - m_Pos + getLastSiblingUnit());
            else
                ((THTreeItem)m_Dad).setWidthUnit(getWidthUnit());
            if(m_Pos > m_Dad.getPos())
                m_Dad.setPos(m_Pos);
            if(m_Pos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - m_Pos);
        }
    
    }

    @Override
    public void setPositionsOfItems(){
        if(isFirstChild()){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = _maxPositions[firstChildLevel] + 1.0;
            m_FirstChild.setLevel(firstChildLevel);
            m_FirstChild.setPos(firstChildPos);

            m_FirstChild.setPositionsOfItems();
        }

        if(isFirstSibling()){
            short firstSiblingLevel = m_Level;
            double firstSiblingNum = m_Pos + getWidthUnit();
            m_FirstSibling.setLevel(firstSiblingLevel);
            m_FirstSibling.setPos(firstSiblingNum);
            m_FirstSibling.setPositionsOfItems();
            setLastSiblingUnit(((THTreeItem)m_FirstSibling).getLastSiblingUnit());
        }else{
            setLastSiblingUnit(getWidthUnit());
        }

        if(m_Dad != null && m_Dad.getFirstChild().equals(this)){
            if(isFirstSibling())
                ((THTreeItem)m_Dad).setWidthUnit(_maxPositions[m_Level] - m_Pos + getLastSiblingUnit());
            else
                ((THTreeItem)m_Dad).setWidthUnit(getWidthUnit());
            if(m_Pos > m_Dad.getPos())
                m_Dad.setPos(m_Pos);
            if(m_Pos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - m_Pos);
        }
        //XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
        //xText.setString(getWidthUnit() + " : " + (int)getLastSiblingUnit());
    }

    // set _horUnit, _shapeWidth, , horSpace, _verUnit, _shapeHeight, verSpace, _groupPosX, _groupPosY
    @Override
    public void setProps(){
        int baseShapeWidth = getDiagramTree().getControlShapeSize().Width;
        int baseShapeHeight = getDiagramTree().getControlShapeSize().Height;

        _shapeWidth =  (int)(baseShapeWidth / (_maxPos + 1));
        _shapeHeight = baseShapeHeight / (_maxLevel + 1);

        _groupPosX = getDiagramTree().getControlShapePos().X;
        _groupPosY = getDiagramTree().getControlShapePos().Y;
    }

    @Override
    public void setPosOfRect(){
        int xCoord = _groupPosX + (int)(_shapeWidth * getPos());
        int yCoord = _groupPosY + _shapeHeight * getLevel();
        setPosition(new Point(xCoord, yCoord));
        setSize(new Size((int)(_shapeWidth * m_WidthUnit), _shapeHeight));
    }

/*
    public THTreeItem getLastSibling(){
        if(isFirstSibling())
            return getLastSibling();
        else
            return this;
    }

    public static void printPos(){
        for(int i=0; i<_maxPositions.length; i++)
            System.out.print(_maxPositions[i] + " ");
        System.out.println();
    }
*/

}
