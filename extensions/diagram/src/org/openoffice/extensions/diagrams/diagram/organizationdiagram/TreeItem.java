package org.openoffice.extensions.diagrams.diagram.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.drawing.XShape;


// TreeItems represent the rectangles of the diagram
public class TreeItem{

    private DiagramTree             m_DiagramTree           = null;
    private XShape                  m_xRectangleShape       = null;
    private int                     m_ID                    = 0;
    private String                  m_sRectangleName        = "";

    private TreeItem                m_Dad                   = null;
    private TreeItem                m_FirstChild            = null;
    private TreeItem                m_FirstSibling          = null;

    private short                   m_Level                 = -1;
    private short                   m_PosNum                = -1;

    protected static short          _maxLevel               = -1;

    protected static short          _level1MaxPos           = -1;
    protected static short          _level2MaxPos           = -1;
    protected static short          _maxPos                 = -1;

    protected static int            _borderTop              = 0;
    protected static int            _borderLeft             = 0;
    protected static int            _groupSizeWidth         = 0;
    protected static int            _groupSizeHeight        = 0;
    protected static int            _horSpace               = 0;
    protected static int            _verSpace               = 0;
    protected static int            _shapeWidth             = 0;
    protected static int            _shapeHeight            = 0;
    protected static int            _horUnit                = 0;

    
    public TreeItem(DiagramTree diagramTree, XShape xShape, TreeItem dad, short level, short num){
        m_DiagramTree = diagramTree;
        m_xRectangleShape = xShape;
        m_ID = getDiagramTree().getShapeID(m_xRectangleShape);
        m_sRectangleName = getDiagramTree().getOrganigram().getShapeName(xShape);
        m_Dad = dad;
        setLevel(level);
        setPosNum(num);
    }

    public DiagramTree getDiagramTree(){
        return m_DiagramTree;
    }

    public void setDad(TreeItem dad){
        m_Dad = dad;
    }

    public TreeItem getDad(){
        return m_Dad;
    }

    public void setFirstSibling(TreeItem treeItem){
        m_FirstSibling = treeItem;
    }

    public TreeItem getFirstSibling(){
        return m_FirstSibling;
    }

    public TreeItem getFirstChild(){
        return m_FirstChild;
    }

    public void setFirstChild(TreeItem treeItem){
        m_FirstChild = treeItem;
    }

    public TreeItem getPreviousSibling(TreeItem treeItem){
        TreeItem previousSibling = null;
        TreeItem item = null;

        if(m_FirstChild != null){
           item = m_FirstChild.getPreviousSibling(treeItem);
           if(item != null)
               previousSibling = item;
        }
        if(getFirstSibling() == treeItem)
            previousSibling = this;

        if(m_FirstSibling != null){
            item = m_FirstSibling.getPreviousSibling(treeItem);
            if(item != null)
               previousSibling = item;
        }
        return previousSibling;
    }

    public void setPosNum(short pos){

        m_PosNum = pos;
        //XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
        //xText.setString(m_Level +":"+m_PosNum);
        if(m_Level == 1)
           if(m_PosNum >_level1MaxPos)
                _level1MaxPos = m_PosNum;
        if(m_Level == 2)
            if(m_PosNum >_level2MaxPos)
                _level2MaxPos = m_PosNum;
        if(_maxPos < m_PosNum)
            _maxPos = m_PosNum;
    }

    public short getPosNum(){
        return m_PosNum;
    }

    public void setLevel(short level){
        m_Level = level;
        if(m_Level > _maxLevel)
            _maxLevel = m_Level;
    }

    public short getLevel(){
        return m_Level;
    }

    public XShape getRectangleShape(){
        return m_xRectangleShape;
    }

    public Size getSize(){
        return m_xRectangleShape.getSize();
    }

    public void setSize(Size size){
        try {
            m_xRectangleShape.setSize(size);
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in TreeItem.setSize(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public Point getPosition(){
        return m_xRectangleShape.getPosition();
    }

    public void setPosition(Point point){
        m_xRectangleShape.setPosition(point);
    }

    // set the treeItems in the tree (set dad-firstChild-firstSibling relations of TreeItems)
    // and set m_level and m_PosNum
    public void initTreeItems(){

        XShape xFirstChildShape = getDiagramTree().getFirstChildShape(m_xRectangleShape);
        if(xFirstChildShape != null){
            short firstChildLevel = (short)(m_Level + 1);
            short firstChildNum = (short)0;
            if(firstChildLevel == 1)
                if(_level1MaxPos > -1)
                    firstChildNum = (short)(_level1MaxPos + 2);
            if(firstChildLevel == 2){
                if(_level2MaxPos == -1)
                    firstChildNum = _level1MaxPos;
                if(_level2MaxPos > -1)
                    firstChildNum = (short)(_level2MaxPos + 2);
            }
            if(firstChildLevel == 3){
                firstChildNum = (short)(_level2MaxPos + 1);
            }
            if(firstChildLevel > 3){
                firstChildNum = (short)(++_level2MaxPos + 1);
            }
            m_FirstChild = new TreeItem(getDiagramTree(), xFirstChildShape, this, firstChildLevel , firstChildNum);
            m_FirstChild.initTreeItems();
        }

        if(m_FirstChild != null){
            TreeItem lastChildItem = getDiagramTree().getTreeItem(getDiagramTree().getLastChildShape(m_xRectangleShape));
            short newPos = (short)(  ( m_FirstChild.getPosNum() + lastChildItem.getPosNum() ) / 2 );
            if(m_Level == 0)
                setPosNum( newPos );
            if(m_Level == 1) {
                if(newPos > getPosNum()){
                    setPosNum( newPos );
                }else if(newPos < getPosNum()){
                    short diff = (short)(getPosNum() - newPos);
                    increaseDescendantsPosNum(diff);
                }
            }
        }

        XShape xFirstSiblingShape = getDiagramTree().getFirstSiblingShape(m_xRectangleShape, m_Dad);
        if(xFirstSiblingShape != null){
            short firstSiblingLevel = m_Level;
            short firstSiblingNum = (short)0;
            if(m_Level == 1 || m_Level == 2)
                firstSiblingNum = (short)(getPosNum() + 2);
            if(m_Level == 2){
                int deep = getDeepOfTreeBranch(this);
                if(deep > 1)
                    firstSiblingNum = (short)(getPosNum() + deep + 1);
            }
            if(m_Level > 2) {
                firstSiblingNum = getPosNum();
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }
            m_FirstSibling = new TreeItem(getDiagramTree(), xFirstSiblingShape, m_Dad, firstSiblingLevel , firstSiblingNum);
            m_FirstSibling.initTreeItems();
        }

    }

    public void printTree(){
        if(m_FirstChild != null)
            m_FirstChild.printTree();
        System.out.println(this);
        if(m_FirstSibling != null)
            m_FirstSibling.printTree();
    }

    public void setPositionsOfItems(){

        if(m_FirstChild != null){
            short firstChildLevel = (short)(m_Level + 1);
            short firstChildNum = (short)0;
            if(firstChildLevel == 1)
                if(_level1MaxPos > -1)
                    firstChildNum = (short)(_level1MaxPos + 2);
            if(firstChildLevel == 2){
                if(_level2MaxPos == -1)
                    firstChildNum = _level1MaxPos;
                if(_level2MaxPos > -1)
                    firstChildNum = (short)(_level2MaxPos + 2);
            }
            if(firstChildLevel == 3)
                firstChildNum = (short)(_level2MaxPos + 1);
            if(firstChildLevel > 3)
                firstChildNum = (short)(++_level2MaxPos + 1);

            m_FirstChild.setLevel(firstChildLevel);
            m_FirstChild.setPosNum(firstChildNum);

            m_FirstChild.setPositionsOfItems();

            TreeItem lastChildItem = getDiagramTree().getTreeItem(getDiagramTree().getLastChildShape(m_xRectangleShape));
            short newPos = (short)(  ( m_FirstChild.getPosNum() + lastChildItem.getPosNum() ) / 2 );
            if(m_Level == 0)
                setPosNum( newPos );
            if(m_Level == 1) {
                if(newPos > getPosNum()){
                    setPosNum( newPos );
                }else if(newPos < getPosNum()){
                    short diff = (short)(getPosNum() - newPos);
                    increaseDescendantsPosNum(diff);
                }
            }
        }
        
        if(m_FirstSibling != null){
 
            short firstSiblingLevel = m_Level;
            short firstSiblingNum = (short)0;
            if(m_Level == 1 || m_Level == 2)
                firstSiblingNum = (short)(getPosNum() + 2);
            if(m_Level == 2){
                int deep = getDeepOfTreeBranch(this);
                if(deep > 1){
                    firstSiblingNum = (short)(getPosNum() + deep + 1);
                }
            }
            if(m_Level > 2) {
                firstSiblingNum = getPosNum();
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }

            m_FirstSibling.setLevel(firstSiblingLevel);
            m_FirstSibling.setPosNum(firstSiblingNum);

            m_FirstSibling.setPositionsOfItems();

        }

    }

    public void searchItem(XShape xShape){
        if(m_FirstChild != null)
            m_FirstChild.searchItem(xShape);
        if(xShape.equals(m_xRectangleShape))
            getDiagramTree().setSelectedItem(this);
        if(m_FirstSibling != null)
            m_FirstSibling.searchItem(xShape);
    }

    public void removeItems(){

        if(m_FirstChild != null)
            m_FirstChild.removeItems();

        XShape xConnShape = getDiagramTree().getDadConnectorShape(m_xRectangleShape);
        if(xConnShape != null){
            getDiagramTree().removeFromConnectors(xConnShape);
            getDiagramTree().getOrganigram().removeShapeFromGroup(xConnShape);
        }
        getDiagramTree().removeFromRectangles(m_xRectangleShape);
        getDiagramTree().getOrganigram().removeShapeFromGroup(m_xRectangleShape);
        
        if(m_FirstSibling != null)
            m_FirstSibling.removeItems();
    }

    public void increaseDescendantsPosNum(short diff){
        m_FirstChild.increasePosNumInBranch(diff);
    }

    public void increasePosNumInBranch(short n){
        if(m_FirstChild != null)
            m_FirstChild.increasePosNumInBranch(n);
        setPosNum((short)(getPosNum() + n));
        if(m_FirstSibling != null)
            m_FirstSibling.increasePosNumInBranch(n);
    }

    public int getDeepOfTreeBranch(TreeItem treeItem){
        if(treeItem.m_FirstChild == null){
            return 0;
        }else{
            int x = 0;
            TreeItem item = treeItem.m_FirstChild;
            while(item != null){
                int y = getDeepOfTreeBranch(item);
                if( y > x )
                    x = y;
                item = item.m_FirstSibling;
            }
            return (x + 1);
        }
    }

    public short getNumberOfItemsInBranch(TreeItem treeItem){
        if(treeItem.m_FirstChild == null){
            return 1;
        }else{
            short n = 1;
            TreeItem item = treeItem.m_FirstChild;
            while(item != null){
                n += getNumberOfItemsInBranch(item);
                item = item.m_FirstSibling;
            }
            return n;
        }
    }

    // set _horSpace, _shapeWidth, _horUnit, _verSpace, _shapeHeight and controlShape size
    public void setProps(){

        Size size = getDiagramTree().getControlShapeSize();
        Point point = getDiagramTree().getControlShapePos();
        _groupSizeWidth = size.Width;
        _groupSizeHeight = size.Height;
        _borderLeft = point.X;
        _borderTop =  point.Y;

        if(_maxPos > 0){
            _horSpace =  _groupSizeWidth * 2 * getDiagramTree().getOrganigram().HORSPACE / ( _maxPos * (getDiagramTree().getOrganigram().WIDTH + getDiagramTree().getOrganigram().HORSPACE) + 2 * getDiagramTree().getOrganigram().WIDTH );
            _shapeWidth = _horSpace * getDiagramTree().getOrganigram().WIDTH / getDiagramTree().getOrganigram().HORSPACE;
        }else{
            _horSpace = 0;
            _shapeWidth = _groupSizeWidth;
        }

        _horUnit = (_shapeWidth + _horSpace ) / 2;

        int verBaseUnit = _groupSizeHeight / ( (_maxLevel+1)*getDiagramTree().getOrganigram().HEIGHT + _maxLevel*getDiagramTree().getOrganigram().VERSPACE );
        if(_maxLevel > 0)
            _verSpace = verBaseUnit * getDiagramTree().getOrganigram().VERSPACE;
        else
            _verSpace = 0;

        _shapeHeight = verBaseUnit * getDiagramTree().getOrganigram().HEIGHT;
    }

    public void display(){
        if(m_FirstChild != null)
            m_FirstChild.display();
        setOptimicPosOfRect();
        if(m_FirstSibling != null)
            m_FirstSibling.display();
    }

    public void setOptimicPosOfRect(){

        int shapeWidth  = _shapeWidth;
        int shapeHeight = _shapeHeight;
        int xCoord = _borderLeft + getPosNum() * _horUnit;//(shapeWidth + _horSpace) / 2;
        int yCoord = _borderTop + m_Level * ( shapeHeight + _verSpace );

        setPosition(new Point(xCoord, yCoord));
        setSize(new Size(shapeWidth, shapeHeight));
    }

    @Override
    public String toString(){
        return "toString: " + m_sRectangleName + " id: " + m_ID + " level: " + m_Level + " num: " + getPosNum();
    }

}
