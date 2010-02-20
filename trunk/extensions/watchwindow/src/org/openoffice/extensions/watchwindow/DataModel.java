package org.openoffice.extensions.watchwindow;

import com.sun.star.container.XNamed;
import com.sun.star.sheet.FormulaToken;
import com.sun.star.sheet.SingleReference;
import com.sun.star.sheet.XFormulaParser;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.table.CellAddress;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import java.util.List;


public class DataModel {
    
    private Controller m_Controller;
    private List<WatchedCell>  m_list; 
    
    public DataModel(Controller controller){
        m_Controller = controller;
        m_list = new ArrayList<WatchedCell>();
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    public void addToDataList(XSpreadsheet xValidSheet, FormulaToken selectedCell) {
        boolean bool = true;
        short num = 0;

        XNamed xNamed = (XNamed)UnoRuntime.queryInterface(XNamed.class, xValidSheet);
        String sheetName = xNamed.getName();

        SingleReference ref = (SingleReference)selectedCell.Data;
        ref.Flags = 0;
        FormulaToken[] tokens = { selectedCell };
        XFormulaParser xParser = m_Controller.getFormulaParser();
        String selectedCellName = xParser.printFormula(tokens, new CellAddress());

        for(WatchedCell item: m_list){
            if(item.getSheetName().equals(sheetName) && item.getCellName().equals(selectedCellName)){
                bool = false;
                num = item.getNum();
                if(num == (short)-1){
                    item.setNum( getController().getNumer());
                    getController().increaseNumer();
                    getController().addToListBox( item.toString(), item.getNum());
                }
            }
         }
         if(bool){
            m_list.add(new WatchedCell(getController(), xValidSheet, sheetName, selectedCell, selectedCellName, getController().getNumer()));
            getController().increaseNumer();
        }
    }
    public void refreshList(){
        for(WatchedCell item: m_list){
            item.refresh();
        }
    }
    public void removeFromDataList(int i){
        //change id of removed element ( m_num = -1 )
        // decrease the next element in the list ( m_num-- )
        short num = 0;
        for(WatchedCell item: m_list){
            num = item.getNum();
            if(num == i){
                item.setNum((short)-1);
            }
            if(num > i)
                item.setNum(--num);
        }
    }
    
    public void removeSheetItemsFromDataList( String sheetName) { 
        short db = 0;
        short num = 0;
        for(WatchedCell item: m_list){
            num = item.getNum();
            if(num > -1){
                if(item.getSheetName().equals(sheetName)){
                    getController().removeFromListBox(num, (short)1);
                    item.setNum((short)-1);
                    getController().decreaseNumer();
                    db++;
                }else{
                    num -= db;
                    item.setNum(num); 
                    getController().removeFromListBox(item.getNum(), (short)1);
                }
            }
        }
    }
    
} 