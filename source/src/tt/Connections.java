package tt;

import java.awt.Color;
import java.awt.Component;
import java.awt.Desktop;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException; 
import java.io.InputStream;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Properties;
import javax.swing.AbstractAction;
import javax.swing.BorderFactory;
import javax.swing.DefaultCellEditor;
import javax.swing.DefaultComboBoxModel;
import javax.swing.DefaultListCellRenderer;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JOptionPane;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.KeyStroke;
import javax.swing.border.Border;
import javax.swing.border.CompoundBorder;
import javax.swing.border.EmptyBorder;
import javax.swing.border.EtchedBorder;
import javax.swing.border.LineBorder;
import javax.swing.event.PopupMenuEvent;
import javax.swing.event.PopupMenuListener;
import javax.swing.filechooser.FileFilter;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Connections {
  int cbi=-1;
  int cbrend=0;

  boolean xx=false;
  boolean arrowPressed=false;
  boolean f4Pressed=false;
  boolean cbRenderSelectionChanged;

  String cbRenderItem;

  // <editor-fold defaultstate="collapsed" desc="variables">
  static final int SUB_IDS=100;                                                                                                // max subjects ids count
  Font tabFont=new Font("Tahoma",Font.PLAIN,16);
  Font headFont=new Font("Trebuchet MS",Font.BOLD,14);
  String text;
  int tabClick=0;
  int prevSelRow, prevSelCol, canEdit;
  boolean keyboardCBChange=false;
  boolean cbItemSelectedWithMouse=false;
  boolean cbPageDownPressed=false;
  boolean cbPageUpPressed=false;
  boolean popupVisible=false;
  boolean cbPageDownUpBeforePopup=false;
  int cbSelIdx;
  int cbPageUpDownIdx=-1;
  JComboBox renderCB;
  JTextField renderTF;
  ItemListener itemListener;
  Connection con;
  ResultSet rs, rsc;
  static HSSFWorkbook w1;                                                                       // excel book and sheet
  static HSSFSheet s1;
  MainForm mf=MainForm.mf;
  JList lsGroups=mf.lsGroups;
  JTable tb=mf.tbData;
  String sqlQuery;                                                                                                           // main query
  String sqlGroupsCapacity;
  boolean newGroupAdded;
  boolean dbLoaded;
  boolean dbSaved;
  boolean rowIsFull;
  boolean aGroupDeleted;
  boolean firstRowFilledBefore;
  static int selectedGroupsCount;                                                                                           // in the list
  static int[] selectedGroupsIndexes=new int[10];                                                               // indexes in the list
  static String[] subjectIds=new String[SUB_IDS];                                                               // ids related to the DB 'id' field
  Timetable.Lesson[][] lessons=Timetable.lessons;                                                 // main array of lessons which will be filled with subjects,teachers,rooms ...
  int[] lesCount=Timetable.lesCount;
  int[] groupsCapacity=Timetable.groupsCapacity;// </editor-fold>

  class CBTableEditor extends DefaultCellEditor{
    JComboBox com;
    boolean cellEditingStopped=false;

    public CBTableEditor(){
      super(new JComboBox());
      setClickCountToStart(2);
      itemListener=new ItemListener(){
        public void itemStateChanged(ItemEvent e){
          if(e.getStateChange()==ItemEvent.DESELECTED){
            System.out.println("desel - "+e.getItem());
          }

          if(e.getStateChange()==ItemEvent.SELECTED){
            System.out.println("sel - "+e.getItem());
            try{
                System.out.println("ckp\n");
                cellKeyPress(com);
                cbItemSelectedWithMouse=true;
              }catch(Exception e1){
              e1.printStackTrace();
            }
          }
        }
      };

      Component jcb=getComponent();
      if(jcb instanceof JComboBox){
        final JComboBox cb=(JComboBox)jcb;

        JTextField tf=(JTextField)cb.getEditor().getEditorComponent();
        cb.setRenderer(new CBRenderer());

        cb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER,0),"none");
        cb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_F4,0),"none");
        cb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_PAGE_DOWN,0),"none");
        cb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_PAGE_UP,0),"none");

        cb.addPopupMenuListener(new PopupMenuListener() {
          @Override
          public void popupMenuWillBecomeVisible(PopupMenuEvent e){
            cbi=cb.getSelectedIndex();
            if(cbi==-1)
              cbi=0;
            cb.addItemListener(itemListener);
            popupVisible=true;
          }

          @Override
          public void popupMenuWillBecomeInvisible(PopupMenuEvent e){
            cbPageUpDownIdx=-1;
          }

          @Override
          public void popupMenuCanceled(PopupMenuEvent e){
          }
        });
        tf.addKeyListener(new KeyAdapter(){
          @Override
          public void keyPressed(KeyEvent e){
            int key=e.getKeyCode();
            JTextField src=(JTextField)e.getSource();
            JComboBox par=(JComboBox)(src).getParent();

            if(key==KeyEvent.VK_F4){
              f4Pressed=true;
              if(!par.isPopupVisible())
                par.showPopup();
              else
                par.hidePopup();
            }
            if(key==10){
              try{
                cellKeyPress(cb);
              }catch(Exception e1){
                e1.printStackTrace();
              }
            }
            if(key==KeyEvent.VK_DOWN|key==KeyEvent.VK_UP){
              arrowPressed=true;
            }
          }
        });
      }
    }

    public Component getTableCellEditorComponent(JTable table,Object value,boolean isSelected,int row,int column){
      com=(JComboBox)super.getTableCellEditorComponent(table,value,isSelected,row,column);

      ResultSet rs=null;
      String tab="";
      ArrayList<String> cdata=new ArrayList<String>();

      boolean useQuery=true;
      boolean formatRoom=false;

      // <editor-fold defaultstate="collapsed" desc="cb-switch">
      switch(column){
        case 1: {
          tab="1Subjects";
          break;
        }
        case 2: {
          useQuery=false;
          cdata.add("18");
          cdata.add("36");
          cdata.add("54");
          break;
        }
        case 3: {
          tab="2Teachers";
          break;
        }
        case 4: {
          tab="3Rooms";
          formatRoom=true;
          break;
        }
        case 5: {
          useQuery=false;
          break;
        }
        case 6: {
          tab="4Groups";
          break;
        }
      }// </editor-fold>
      // <editor-fold defaultstate="collapsed" desc="cb-get">
      if(useQuery){
        rs=execQuery("SELECT name FROM "+tab,true);
        if(rs!=null){
          try{
            while(rs.next()){
              String val=rs.getString(1);
              if(formatRoom){
                if(val.equals("000 1"))
                  val="ñï. çàë";
                else{
                  String ss;
                  if(val.charAt(val.length()-2)==' ')
                    ss=val.substring(0,val.length()-2);
                  else
                    ss=val.substring(0,val.length()-1);
                  val=ss+'('+val.charAt(val.length()-1)+')';
                }
              }
              cdata.add(val);
            }
            com.setModel(new DefaultComboBoxModel(cdata.toArray()));
            com.removeItemListener(itemListener);
            com.setSelectedItem(value);
//            com.addItemListener(itemListener);
          }catch(SQLException e){
            e.printStackTrace();
          }
        }
      }
      else if(cdata.size()!=0){
        com.setModel(new DefaultComboBoxModel(cdata.toArray()));
        com.removeItemListener(itemListener);
        com.setSelectedItem(value);
//        com.addItemListener(itemListener);
      }
      else{
        com.setModel(new DefaultComboBoxModel(new Object[0]));
      }
      // </editor-fold>

      com.setMaximumRowCount(15);
      com.setEditable(true);
      com.setFont(tabFont);

      return com;
    }

    @Override
    public void cancelCellEditing(){
      com.removeItemListener(itemListener);
      super.cancelCellEditing();
    }
  }
  class CBRenderer extends DefaultListCellRenderer{
    JLabel com;
    int pginc=5;
    int x=0;

    public Component getListCellRendererComponent(JList list,Object value,int index,boolean isSelected,boolean cellHasFocus){
      com=(JLabel)super.getListCellRendererComponent(list,value,index,isSelected,cellHasFocus);
      com.setText((String)value);

      if(popupVisible){
        popupVisible=false;
        index=cbi;
        int id=index;
        list.ensureIndexIsVisible(id);
        list.setSelectedIndex(id);

        if(cbPageDownUpBeforePopup)
          cbPageDownUpBeforePopup=false;
        else
          cbPageUpDownIdx=0;
      }

      if(cbPageDownPressed){
        if(cbPageDownUpBeforePopup)
          cbPageDownUpBeforePopup=false;
        cbPageDownPressed=false;

        System.out.println("cbi1-dw:"+cbi);
        if(cbPageUpDownIdx==-1){
          cbi=index;
          cbPageUpDownIdx=0;
        }
        else if(cbi==list.getModel().getSize()-1)
          cbi=0;
        else if(cbi+pginc>list.getModel().getSize()-1)
          cbi=list.getModel().getSize()-1;
        else
          cbi+=pginc;
        System.out.println("cbi2-dw:"+cbi);
        System.out.println("");

        list.ensureIndexIsVisible(cbi);
        list.setSelectedIndex(cbi);
      }
      else if(cbPageUpPressed){
        if(cbPageDownUpBeforePopup)
          cbPageDownUpBeforePopup=false;
        cbPageUpPressed=false;

        System.out.println("cbi1-up:"+cbi);
        if(cbPageUpDownIdx==-1){
          cbi=index;
          cbPageUpDownIdx=0;
        }
        else if(cbi==0)
          cbi=list.getModel().getSize()-1;
        else if(cbi-pginc<0)
          cbi=0;
        else
          cbi-=pginc;
        System.out.println("cbi2-up:"+cbi);
        System.out.println("");

        list.ensureIndexIsVisible(cbi);
        list.setSelectedIndex(cbi);
      }

      if(isSelected){
        if(arrowPressed){
          arrowPressed=false;
          cbi=index;
        }
        if(f4Pressed){
          f4Pressed=false;
          cbi=index;
        }

        keyboardCBChange=true;
        cbSelIdx=index;

        setBackground(Color.gray);
        setForeground(Color.white);
      }
      else{
        int gray=240;
        setBackground(new Color(gray,gray,gray));
        setForeground(Color.black);
      }

      return com;
    }
  }
  
  class TFTableEditor extends DefaultCellEditor{
    JTextField tf;
    boolean lock=false;

    public TFTableEditor(){
      super(new JTextField());
      getComponent().addKeyListener(new KeyAdapter(){
        @Override
        public void keyPressed(KeyEvent e){
          int key=e.getKeyCode();
          if(key==KeyEvent.VK_ENTER)
            try{
              cellKeyPress(tf);
            }catch(Exception e1){
              e1.printStackTrace();
            }
          if(key==KeyEvent.VK_ESCAPE){
            cancelCellEditing();
          }
        }
      });
      getComponent().addFocusListener(new FocusAdapter(){
        public void focusGained(FocusEvent e){
//          tx.selectAll();
        }
      });
      setClickCountToStart(2);
    }

    public Component getTableCellEditorComponent(JTable table,Object value,boolean isSelected,int row,int column){
      tf=(JTextField)super.getTableCellEditorComponent(table,value,isSelected,row,column);
      Border b1=new CompoundBorder(new EtchedBorder(0,new Color(150,200,240),Color.black),
              new EtchedBorder(1,new Color(150,200,240),Color.darkGray));
      Border b2=BorderFactory.createEmptyBorder(0,2,0,0);
      tf.setBorder(new CompoundBorder(b1,b2));

      text=(String)value;
      tf.setText(text);
      tf.setFont(tabFont);

      return tf;
    }
  }
  class TableRenderer extends DefaultTableCellRenderer{
    public Component getTableCellRendererComponent(JTable table,Object value,boolean isSelected,boolean hasFocus,int row,int column){
      JLabel lb=(JLabel)super.getTableCellRendererComponent(table,value,isSelected,hasFocus,row,column);

      table.setCellSelectionEnabled(true);
      if(table.isCellSelected(row,column)){
        if(cbItemSelectedWithMouse){
          if(tb.editCellAt(row,column))
            tb.getEditorComponent().requestFocus();
          cbItemSelectedWithMouse=false;
        }

        int clr=230;
        setBackground(new Color(clr,clr,clr));
        setForeground(Color.black);

        Border b1=new CompoundBorder(new EmptyBorder(1,1,1,1),new LineBorder(new Color(30,120,170),2));
        Border b2=new EmptyBorder(0,3,0,0);
        lb.setBorder(new CompoundBorder(b1,b2));
      }
      else{
        setBackground(Color.white);
        setForeground(Color.black);

        Border b1=BorderFactory.createMatteBorder(1,1,0,0,Color.gray);
        Border b2=BorderFactory.createEmptyBorder(0,5,0,0);
        lb.setBorder(new CompoundBorder(b1,b2));
      }

      return lb;
    }
  }
  class TableModel extends DefaultTableModel{
    public TableModel(Object[][] data,Object[] columnNames){
      super(data,columnNames);
    }

    @Override
    public boolean isCellEditable(int row,int column){
      if(column==0)
        return false;
      else
        return super.isCellEditable(row,column);
    }
  }
  class FileChooseFilter extends FileFilter{
    public boolean accept(File f){
      if(f.isDirectory()){
        return true;
      }
      else{
        String ext=f.getName();
        int dotIndex=ext.lastIndexOf('.');
        if(dotIndex>-1)
          ext=ext.substring(dotIndex);

        if(ext.equals(".mdb"))
          return true;
        else
          return false;
      }
    }

    public String getDescription(){
      return ".mdb files";
    }
  }

  void initTable(){
    // <editor-fold defaultstate="collapsed" desc="inits">
    tb.addKeyListener(new KeyAdapter(){
      @Override
      public void keyPressed(KeyEvent e){
        int key=e.getKeyCode();
        boolean cond=(key>=KeyEvent.VK_COMMA&&key<=KeyEvent.VK_DIVIDE)||(key>=KeyEvent.VK_BACK_QUOTE&&key<=KeyEvent.VK_QUOTE)||(key>=KeyEvent.VK_DEAD_GRAVE&&key<=KeyEvent.VK_UNDERSCORE)||key==KeyEvent.VK_BACK_SPACE||key==KeyEvent.VK_SPACE||key==KeyEvent.VK_ENTER;
        if(cond){
          int r=tb.getSelectedRow();
          int c=tb.getSelectedColumn();

          if(tb.editCellAt(r,c)){
            tb.getEditorComponent().requestFocus();
          }
        }
      }
    });
    tb.addMouseListener(new MouseAdapter(){
      public void mousePressed(MouseEvent e){
//        int r=tb.getSelectedRow();
//        int c=tb.getSelectedColumn();
//
//        if(prevSelRow!=r||prevSelCol!=c){
//          tb.changeSelection(r,c,false,false);
//          canEdit=1;
//        }
//        else{
//          tb.getDefaultEditor(Object.class).cancelCellEditing();
//          tb.editCellAt(r,c);
//          tb.getEditorComponent().requestFocus();
//          canEdit=0;
//        }
//
//        prevSelRow=r;
//        prevSelCol=c;
      }
    });

    tb.setDefaultEditor(Object.class,new CBTableEditor());
    tb.setDefaultRenderer(Object.class,new TableRenderer());

    tb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_ENTER,0),"none");
    tb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_F4,0),"none");
    tb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_PAGE_DOWN,0),"pg-down");
    tb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_PAGE_UP,0),"pg-up");
    tb.getInputMap(JTable.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT).put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE,0),"esc");// </editor-fold>

    tb.getActionMap().put("esc",new AbstractAction(){
      @Override
      public void actionPerformed(ActionEvent e){
        System.out.println("esc");
        tb.getDefaultEditor(Object.class).cancelCellEditing();
      }
    });
    tb.getActionMap().put("pg-down",new AbstractAction(){
      @Override
      public void actionPerformed(ActionEvent e){
        int r=tb.getEditingRow();
        int c=tb.getEditingColumn();
        if(r!=-1){
          JComboBox cc=(JComboBox)tb.getEditorComponent();
          cbPageDownUpBeforePopup=true;
          if(!cc.isPopupVisible())
            cc.setPopupVisible(true);
          else
            cc.repaint();
          cbPageDownPressed=true;
        }
      }
    });
    tb.getActionMap().put("pg-up",new AbstractAction(){
      @Override
      public void actionPerformed(ActionEvent e){
        int r=tb.getEditingRow();
        int c=tb.getEditingColumn();
//        System.out.println(r+","+c);
        if(r!=-1){
          JComboBox cc=(JComboBox)tb.getEditorComponent();
          cbPageDownUpBeforePopup=true;
          if(!cc.isPopupVisible())
            cc.setPopupVisible(true);
          else
            cc.repaint();
          cbPageUpPressed=true;
        }
      }
    });

    tb.setRowHeight(30);
  }

  void newTimetable(){
    rowIsFull=false;
    Connections.selectedGroupsCount=0;
    int r=0,c=1;

    for(int i=0;i<SUB_IDS;i++)
      subjectIds[i]="";

    formatTable(new Object[1][0],c);
    tb.setValueAt("1",0,0);

    if(tb.editCellAt(r,c))
      tb.getEditorComponent().requestFocus();
    lsGroups.setListData(new Object[0]);

    String projectPath=System.getProperty("user.dir");
    new File(projectPath+"\\tmp").mkdir();
    String src=projectPath+"\\templates\\tmp1.mdb";
    String dest=projectPath+"\\tmp\\tmp.mdb";

    dbCopy(src,dest);
    dbConnect(dest);

//    fillFirstRow();
  }

  void fillFirstRow(){
    int r=0;
    firstRowFilledBefore=true;

    tb.getDefaultEditor(Object.class).cancelCellEditing();
    
    tb.setValueAt("subject1",0,1);
    tb.setValueAt("18",0,2);
    tb.setValueAt("teacher1",0,3);
    tb.setValueAt("123(1)",0,4);
    tb.setValueAt("60",0,5);
    tb.setValueAt("gr1",0,6);

    for(int i=1;i<7;i++){
      tb.editCellAt(r,i);
      Component comp=tb.getEditorComponent();
      comp.dispatchEvent(new KeyEvent(comp,KeyEvent.KEY_PRESSED,System.currentTimeMillis(),0,KeyEvent.VK_ENTER,'\n'));
    }
  }

  void dbLoad(){
    String projectPath=System.getProperty("user.dir");
    String openPath=projectPath+"\\examples";
    dbLoaded=false;
    newGroupAdded=false;

    JFileChooser fchLoad=new JFileChooser();
    fchLoad.setSelectedFile(new File(openPath+"\\pr.mdb"));
//    fchLoad.setSelectedFile(new File(projectPath+"/projects/timetable.mdb"));
    fchLoad.setPreferredSize(new Dimension(650,400));
    fchLoad.setCurrentDirectory(new File(openPath));
    fchLoad.addChoosableFileFilter(new FileChooseFilter());
    fchLoad.setAcceptAllFileFilterUsed(false);

    if(fchLoad.showOpenDialog(mf)==JFileChooser.APPROVE_OPTION){
      String loadPath=fchLoad.getSelectedFile().getPath();
      String src=loadPath;
      String dest=projectPath+"\\tmp\\tmp.mdb";

      dbCopy(src,dest);
      dbConnect(dest);
      dbLoaded=true;
      rowIsFull=true;
    }
  }

  void getGroupNames(){                                                                                                // get names and show them in the JList component
    if(dbLoaded|newGroupAdded|aGroupDeleted){
      lsGroups.setListData(new Object[0]);
      String query="select [4Groups].name FROM 4Groups;";
      ArrayList<String> grNames=null;

      try{
        ResultSet rsLoc=con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_READ_ONLY).executeQuery(query);
        if(!rsLoc.next())
          lsGroups.setListData(new Object[0]);
        else{
          grNames=new ArrayList<String>();
          rsLoc.beforeFirst();
          while(rsLoc.next())
            grNames.add(rsLoc.getString(1));
          lsGroups.setListData(grNames.toArray());                                                                       // fill the list with groups names
          lsGroups.setSelectionInterval(0,0);
        }
      }catch(SQLException e){
        e.printStackTrace();
      }
    }
  }

  void fillTable(){                                                                                                        // fill JTable on the form
    for(int i=0;i<SUB_IDS-1;i++)
      subjectIds[i]="";
    selectedGroupsCount=0;

    String s1="SELECT [1Subjects].id, [1Subjects].name, [1Subjects].load, [2Teachers].name, "+
                     "[3Rooms].name, [3Rooms].capacity, [4Groups].name ";
    String s2_1="FROM (2Teachers INNER JOIN ((1Subjects INNER JOIN (3Rooms INNER JOIN SubjectRoom ON "+
                         "[3Rooms].name = SubjectRoom.aud_name) ON [1Subjects].id = SubjectRoom.subj_id) "+
                         "INNER JOIN SubjectTeacher ";
    String s2_2="ON [1Subjects].id = SubjectTeacher.subj_id) ON [2Teachers].id = SubjectTeacher.tch_id) "+
                         "INNER JOIN (4Groups INNER JOIN SubjectGroup ON [4Groups].id = SubjectGroup.group_id) ON "+
                         "[1Subjects].id = SubjectGroup.subj_id ";
    String s2=s2_1+s2_2;

    String s3="";
    int[] selIndexes=lsGroups.getSelectedIndices();
    for(int i=0;i<selIndexes.length;i++){						// looking through all of the selected items in the List
      String selValue=(String)lsGroups.getSelectedValues()[i];
      if(selectedGroupsCount==0)                                                    
        s3="([4Groups].name)='"+selValue+"'";
      else                                                                          
        s3=s3+" Or ([4Groups].name)='"+selValue+"'";                   // if there is only 1 group selected
      selectedGroupsIndexes[selectedGroupsCount]=selIndexes[i];                                                             // indexes of selected groups
      selectedGroupsCount++;                                                        
    }                                                                               
    s3="WHERE (("+s3+"))";                                                          
    String s4=" ORDER BY [4Groups].name,[1Subjects].name;";                       

    sqlQuery=s1+s2+s3+s4;                                                                                                           // main query is constructed
    sqlGroupsCapacity="SELECT [4Groups].capacity FROM 4Groups "+s3+";";

//************* Fill the Table **************
    try{
      rs=con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_READ_ONLY).executeQuery(sqlQuery); // scroll_sensitive - to use beforeFirst() method
      int rowCount=0;
      while(rs.next())
        subjectIds[rowCount++]=rs.getString(1);
      rs.beforeFirst();

      int colCount=rs.getMetaData().getColumnCount();
      String[][] tabData=new String[rowCount][colCount];                                                      // main data array

      int i=0;
      while(rs.next()){                                                                                                                // main data retreiving loop
        tabData[i][0]=String.valueOf(i+1);
        for(int j=1;j<colCount;j++)
          if(j==4){
            String tmpRoom=rs.getString(j+1);
            String s=tmpRoom;
            if(s.equals("000 1"))
              tmpRoom="gym";
            else{
              String ss;
              if(s.charAt(s.length()-2)==' ')
                ss=s.substring(0,s.length()-2);
              else
                ss=s.substring(0,s.length()-1);
              tmpRoom=ss+'('+s.charAt(s.length()-1)+')';
            }
            tabData[i][j]=tmpRoom;
          }
          else
            tabData[i][j]=rs.getString(j+1);
        i++;
      }

      formatTable(tabData,1);
    }catch(SQLException e){
      e.printStackTrace();
    }

    tb.getColumnModel().getColumn(5).setCellEditor(new TFTableEditor());
  }

  void importData(){
    Timetable.Lesson[] ptLes=new Timetable.Lesson[10];
    Timetable.Lesson[] hgLes=new Timetable.Lesson[10];

    for(int i=0;i<Timetable.GROUPS;i++)
      for(int j=0;j<Timetable.SUBJECTS;j++)
        lessons[i][j]=null;

    int ptlCount=0;
    int hglCount=0;
    int gr=0;
    boolean editPTL=true;

    for (int i = 0; i <selectedGroupsCount; i++){
      gr=selectedGroupsIndexes[i];
      Timetable.lesCount[gr]=0;
    }

    try {
      rs.beforeFirst();
      while(rs.next()){
        String tmpSubject=rs.getString(2);
        int tmpLoad=rs.getInt(3);
        String tmpTeacher=rs.getString(4);
        String tmpRoom=rs.getString(5);

        String s=tmpRoom;
        if(s.equals("000 1"))                                                                                                       // Physical Training lessons
          tmpRoom="";
        else{
          String ss;
          if(s.charAt(s.length()-2)==' ')                                                                                   // put () around the building number
            ss=s.substring(0,s.length()-2);
          else
            ss=s.substring(0,s.length()-1);
          tmpRoom=ss+'('+s.charAt(s.length()-1)+')';
        }

        int tmpRoomCapacity=rs.getInt(6);
        String tmpGroup=rs.getString(7);

        if(tmpGroup.equals(lsGroups.getModel().getElementAt(0)))
          gr=0;
        else if(tmpGroup.equals(lsGroups.getModel().getElementAt(1))){
          gr=1;                                                                                                                                   // Physical Training lessons are same for all groups
          if(selectedGroupsCount>1)                                                                                              // so if they were imported from one group subjects
            editPTL=false;                                                                                                                // we don't need to import them from others groups subjects
        }

        if(tmpSubject.equals("Physical Training")){                                                               // here we check if PT lessons were already imported from the DB
          if(editPTL==true){
            ptLes[ptlCount]=new Timetable.Lesson();
            ptLes[ptlCount].subject=tmpSubject;
            ptLes[ptlCount].load=tmpLoad;
            ptLes[ptlCount].teacher=tmpTeacher;
            ptLes[ptlCount].room=tmpRoom;
            ptLes[ptlCount].roomCapacity=tmpRoomCapacity;
            ptlCount++;
          }
        }
        else if(tmpSubject.equals("Foreign Language (prof. area)")){
          tmpSubject="Foreign Language";
          hgLes[hglCount]=new Timetable.Lesson();
          hgLes[hglCount].subject=tmpSubject;
          hgLes[hglCount].load=tmpLoad;
          hgLes[hglCount].teacher=tmpTeacher;
          hgLes[hglCount].room=tmpRoom;
          hgLes[hglCount].roomCapacity=tmpRoomCapacity;
          hglCount++;
        }
        else if(tmpLoad==54){                                                                                                       // main subjects (18-36h and not PTs neither HGs)
          tmpLoad=36;                                                                                                                       // if subject is 54h then divide it into 2 subjects (36+18h)
          for (int i = 0; i <2; i++){
            lessons[gr][lesCount[gr]]=new Timetable.Lesson();
            lessons[gr][lesCount[gr]].subject=tmpSubject;
            lessons[gr][lesCount[gr]].load=tmpLoad;
            lessons[gr][lesCount[gr]].teacher=tmpTeacher;
            lessons[gr][lesCount[gr]].room=tmpRoom;
            lessons[gr][lesCount[gr]].roomCapacity=tmpRoomCapacity;

            lesCount[gr]++;
            tmpLoad=18;
          }
        }
        else{
          lessons[gr][lesCount[gr]]=new Timetable.Lesson();
          lessons[gr][lesCount[gr]].subject=tmpSubject;
          lessons[gr][lesCount[gr]].load=tmpLoad;
          lessons[gr][lesCount[gr]].teacher=tmpTeacher;
          lessons[gr][lesCount[gr]].room=tmpRoom;
          lessons[gr][lesCount[gr]].roomCapacity=tmpRoomCapacity;

          lesCount[gr]++;
        }
      }

// *** Physical Training lessons ***
      if(ptlCount>0){
        int pti1,pti2;
        for (int i = 0; i <selectedGroupsCount; i++){
          gr=selectedGroupsIndexes[i];                                                                                        // imported PT lessons must be included to all of the groups selected in the ListBox
          pti1=(int)(Math.random()*ptlCount);
          do {
            pti2=(int)(Math.random()*ptlCount);                                                                         // join 2 random teachers for one Physical Training lesson
          } while (pti1==pti2);
          lessons[gr][lesCount[gr]]=new Timetable.Lesson();
          lessons[gr][lesCount[gr]]=ptLes[pti1].clone();                                                                      // add this PT lesson to the end of the 'lessons' array
          lessons[gr][lesCount[gr]].teacher=lessons[gr][lesCount[gr]].teacher+", "+ptLes[pti2].teacher;
          lesCount[gr]++;
        }
      }

// *** Half Group 18h lessons ***
      gr=1;                                                                                                                                      // such lessons are only in the 2PR group
      if(hglCount>0){
        int hgi1,hgi2;
        hgi1=(int)(Math.random()*hglCount);
        do {
            hgi2=(int)(Math.random()*hglCount);
        } while (hgLes[hgi1].teacher==hgLes[hgi2].teacher);
        lessons[gr][lesCount[gr]]=new Timetable.Lesson();
        lessons[gr][lesCount[gr]]=hgLes[hgi1];
        lesCount[gr]++;

        lessons[gr][lesCount[gr]]=new Timetable.Lesson();
        lessons[gr][lesCount[gr]]=hgLes[hgi2];
        lesCount[gr]++;
      }

//========= Groups capacity ==========
      rsc = con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_READ_ONLY).executeQuery(sqlGroupsCapacity);
      rsc.next();
      for (int i = 0; i <selectedGroupsCount; i++){
        gr=selectedGroupsIndexes[i];
        groupsCapacity[gr]=rsc.getInt(1);
        rsc.next();
      }
    } catch (Exception e) {
      e.printStackTrace();
    }

    if(mf.pnum==2){
      int x=0;
    }
  }

  void openExcel(){
    String projectPath=System.getProperty("user.dir");
    String projectsDir=projectPath+"/projects/";
    File f=new File(projectsDir);
    if(!f.exists())
      f.mkdir();

    f=new File(projectsDir+"timetable.xls");
    if(f.exists()){
      File f1=new File(projectsDir);

      FilenameFilter ff=new FilenameFilter(){
        public boolean accept(File f,String s){
          return s.startsWith("timetable");
        }
      };
      String[] list=f1.list(ff);
      if(list!=null){
        Arrays.sort(list,new Comparator<String>(){
          public int compare(String a,String b){            // sort by the numbers (the default sorting doesn't sort numbers properly)
            int i=a.indexOf('.')-1;
            int j=b.indexOf('.')-1;

            while((a.charAt(i)>=48)&&(a.charAt(i)<=57)) // find the index before the first digit of the number
              i--;
            while((b.charAt(j)>=48)&&(b.charAt(j)<=57))
              j--;
            if(i==a.indexOf('.')-1)// if there is no number - add 0, to avoid errors
              a=a.substring(0,a.indexOf('.'))+'0'+a.substring(a.indexOf('.'));

            int a1=Integer.valueOf(a.substring(i+1,a.indexOf('.'))); // get numbers
            int b1=Integer.valueOf(b.substring(j+1,b.indexOf('.')));

            if(a1>b1)
              return 1;
            else
              return 0;
          }
        });

        String s=list[list.length-1];                                                                                        // the last element
        int i=s.indexOf('.')-1;
        while((s.charAt(i)>=48)&&(s.charAt(i)<=57))                                                              // get the number
          i--;
        if(i==s.indexOf('.')-1)
          s=s.substring(0,s.indexOf('.'))+"0"+s.substring(s.indexOf('.'));                     // if there is only 'timetable.xls' file

        int nx=Integer.valueOf(s.substring(i+1,s.indexOf('.')))+1;                                   // add 1
        s=s.substring(0,i+1)+String.valueOf(nx)+s.substring(s.indexOf('.'));                // new filename
        f=new File(projectsDir+s);
      }
      else
        System.out.println("no");
    }

    try{
      FileOutputStream fout=new FileOutputStream(f);
      w1.write(fout);
      fout.close();
      Desktop.getDesktop().open(f);                                                                                          // open xls and return to the programm execution
    }catch(IOException e){
      e.printStackTrace();
    }
  }

  void dbSave(){
    String projectPath=System.getProperty("user.dir");
    String saveDir=projectPath+"\\projects";
    File f=new File(saveDir);
    if(!f.exists())
      f.mkdir();
    dbSaved=false;

    JFileChooser fchSave=new JFileChooser();
    fchSave.setSelectedFile(new File(projectPath+"\\timetable.mdb"));
    fchSave.setPreferredSize(new Dimension(650,400));
    fchSave.setCurrentDirectory(new File(saveDir));
    fchSave.addChoosableFileFilter(new FileChooseFilter());
    fchSave.setAcceptAllFileFilterUsed(false);

    if(fchSave.showSaveDialog(mf)==JFileChooser.APPROVE_OPTION){
      String savePath=fchSave.getSelectedFile().getPath();
      String src=projectPath+"\\tmp\\tmp.mdb";
      String dest=savePath;
      dbCopy(src,dest);
      dbSaved=true;
    }
  }

  void dbConnect(String dbPath){
    String url="jdbc:odbc:Driver={Microsoft Access Driver (*.mdb)};DBQ="+dbPath;
    Properties props=new Properties();
    props.put("charSet","windows-1251");

    try{
      dbDisconnect(false);
      con=DriverManager.getConnection(url,props);
      
      String dbPath1 = "d:\\Documents\\8-projects\\eclipse\\java\\Timetable\\tmp\\tmp.mdb";
      mf.setTitle(mf.title+"   [loaded database: "+dbPath1+"]");
      
      // mf.setTitle(mf.title+"   [loaded database: "+dbPath+"]");
      
    }catch(SQLException e){
      e.printStackTrace();
    }
  }

  void dbDisconnect(boolean delTemp){
    try{
      if(con!=null){
        con.close();
      }

      if(delTemp){
        String projectPath=System.getProperty("user.dir");
        new File(projectPath+"\\tmp\\tmp.mdb").delete();
        new File(projectPath+"\\tmp").delete();
      }
    }catch(SQLException e){
      e.printStackTrace();
    }
  }

  void xlConnect(){
    w1=new HSSFWorkbook();
    s1=w1.createSheet("Timetable");
  }
  
  void dbCopy(String src,String dest){
    try{
      InputStream in=new FileInputStream(src);
      OutputStream out=new FileOutputStream(dest);
      byte[] buf=new byte[1024];
      int len;

      while((len=in.read(buf))>0)
        out.write(buf,0,len);
      in.close();
      out.close();
    }catch(IOException e){
      e.printStackTrace();
    }
  }

  void formatTable(Object[][] tabData,int selCol){
    String[] tabHead={"N","Subject","Load","Teacher","Room","Capacity","Group"};
    tb.setModel(new TableModel(tabData,tabHead));

    int widths[]={35,430,40,200,90,60,50};
    for(int j=0;j<tb.getColumnCount();j++)
      tb.getColumnModel().getColumn(j).setPreferredWidth(widths[j]);
    tb.setRowHeight(30);

    tb.getColumnModel().getColumn(0).setCellEditor(new TFTableEditor());
    tb.getColumnModel().getColumn(5).setCellEditor(new TFTableEditor());

    tb.getTableHeader().setFont(headFont);
    tb.setFont(tabFont);

    tb.requestFocus();
    tb.changeSelection(0,selCol,false,false);
  }

  void add(){
    int rowNumber;
    if(!rowNotComplete(tb.getSelectedRow())){
      rowIsFull=false;
      ((DefaultTableModel)tb.getModel()).addRow((Object[])null);
      rowNumber=tb.getRowCount();
      tb.setValueAt(String.valueOf(rowNumber),rowNumber-1,0);
//      tb.setValueAt("Áàçè äàíèõ (ëåê.)",rowNumber-1,1);                   //!!! temp
    }
    else
      rowNumber=tb.getSelectedRow()+1;

    if(tb.editCellAt(rowNumber-1,1)){
      tb.getEditorComponent().requestFocus();
      tb.changeSelection(rowNumber-1,1,false,false);
    }

    int r=tb.getEditingRow();
    int c=tb.getEditingColumn();

    int x=0;
  }

  void delete(){
    if(lsGroups.getSelectedIndex()==-1){
      formatTable(new Object[1][0],0);
      tb.setValueAt("1",0,0);

      tb.changeSelection(0,1,false,false);
      if(tb.editCellAt(0,1))
        tb.getEditorComponent().requestFocus();
    }
    else{
      int delRow=tb.getEditingRow();
      if(delRow==-1)
        delRow=tb.getSelectedRow();

      if(!rowNotComplete(delRow)){
        execQuery("DELETE FROM SubjectTeacher WHERE SubjectTeacher.subj_id="+subjectIds[delRow]);
        execQuery("DELETE FROM SubjectRoom WHERE SubjectRoom.subj_id="+subjectIds[delRow]);
        execQuery("DELETE FROM SubjectGroup WHERE SubjectGroup.subj_id="+subjectIds[delRow]);
      }

      if(tb.getRowCount()==1){
        String delGroup=(String)lsGroups.getSelectedValue();
        execQuery("DELETE FROM 4Groups WHERE name='"+delGroup+"'");
        aGroupDeleted=true;
        getGroupNames();
      }

      refresh();
      if(delRow==tb.getRowCount())
        tb.changeSelection(delRow-1,1,false,false);
      else
        tb.changeSelection(delRow,1,false,false);
    }
  }

  void refresh(){
    if(lsGroups.getSelectedIndex()==-1){
      delete();
    }
    else{
      int[] sel=lsGroups.getSelectedIndices();
      lsGroups.clearSelection();
      lsGroups.setSelectionInterval(sel[0],sel[sel.length-1]);
    }
  }

  void cellKeyPress(JComponent comp) throws SQLException{
    boolean isTextField=false;
    String editValue=null;

    JComboBox cb=null;
    JTextField tf=null;

    if(comp instanceof JTextField){
      isTextField=true;
      tf=(JTextField)comp;
      editValue=(String)tf.getText();
    }
    else{
      cb=(JComboBox)comp;
      editValue=(String)cb.getEditor().getItem();
    }

    if(keyboardCBChange&cb!=null){
      cb.setSelectedIndex(cbSelIdx);
      keyboardCBChange=false;
      editValue=(String)cb.getEditor().getItem();
    }

    boolean selectNextCell;
    ResultSet rs;

    selectNextCell=true;
    newGroupAdded=false;

    int r=tb.getEditingRow();
    int c=tb.getEditingColumn();

    if(r==-1)
      r=tb.getSelectedRow();
    if(c==-1)
      c=tb.getSelectedColumn();

    if(editValue.length()!=0){
      String sid=subjectIds[r];

      rowIsFull=true;
      if(rowNotComplete(r))
        rowIsFull=false;

      if(firstRowFilledBefore)
        rowIsFull=false;                        //-----------------temp-----

      switch(c){
        // <editor-fold defaultstate="collapsed" desc="1-subject">
        case 1: {
          String tmpSubject=editValue;
          String tmpSID;
          boolean addNewSubject;

          if(sid==null)
            addNewSubject=true;
          else if(sid.length()==0)
            addNewSubject=true;
          else
            addNewSubject=false;

          if(addNewSubject){
            rs=execQuery("SELECT Max([1Subjects].id) FROM 1Subjects");
            rs.next();
            int intSID=rs.getInt(1);
            if(intSID==0)
              tmpSID="1001";
            else
              tmpSID=String.valueOf(intSID+1);
            execQuery("INSERT INTO 1Subjects (id,name) VALUES ("+tmpSID+",'"+tmpSubject+"')");
            subjectIds[r]=tmpSID;
          }
          else{
            execQuery("UPDATE 1Subjects SET [1Subjects].name='"+tmpSubject+"' WHERE ((([1Subjects].id)="+sid+"))");
            refresh();
          }
          break;
        }// </editor-fold>
        // <editor-fold defaultstate="collapsed" desc="2-load">
        case 2: {
          String tmpLoad=editValue;
          boolean wrongLoadType=false;

          for(int i=0;i<tmpLoad.length();i++)
            if(!(tmpLoad.charAt(i)>='0'&&tmpLoad.charAt(i)<='9')){
              wrongLoadType=true;
              break;
            }

          if(wrongLoadType){
            cb.getEditor().selectAll();
            JOptionPane.showMessageDialog(mf,"Enter a number","Wrong type",JOptionPane.ERROR_MESSAGE);
            selectNextCell=false;
          }
          else{
            execQuery("UPDATE 1Subjects SET [1Subjects].load="+tmpLoad+" WHERE ((([1Subjects].id)="+sid+"))");
            if(rowIsFull)
              refresh();
          }
          break;
        }// </editor-fold>
        // <editor-fold defaultstate="collapsed" desc="3-teacher">
        case 3: {
          String tmpTeacher=editValue;
          String tmpTID;

          rs=execQuery("SELECT [2Teachers].id FROM 2Teachers WHERE ((([2Teachers].name)='"+tmpTeacher+"'))");
          if(!rs.next()){
            rs=execQuery("SELECT Max([2Teachers].id) FROM 2Teachers");
            rs.next();
            int intTID=rs.getInt(1);
            if(intTID==0)
              tmpTID="2001";
            else
              tmpTID=String.valueOf(intTID+1);
            execQuery("INSERT INTO 2Teachers (id,name) VALUES ("+tmpTID+",'"+tmpTeacher+"')");

            rs=execQuery("SELECT SubjectTeacher.subj_id FROM SubjectTeacher WHERE (((SubjectTeacher.subj_id)="+sid+"))");
            if(!rs.next())
              execQuery("INSERT INTO SubjectTeacher (subj_id,tch_id) VALUES ("+sid+","+tmpTID+")");
            else
              execQuery("UPDATE SubjectTeacher SET SubjectTeacher.tch_id = "+tmpTID+
                      " WHERE (((SubjectTeacher.subj_id)="+sid+"))");
          }
          else{
            tmpTID=rs.getString(1);
            rs=execQuery("SELECT SubjectTeacher.subj_id FROM SubjectTeacher WHERE (((SubjectTeacher.subj_id)="+sid+"))");
            if(!rs.next())
              execQuery("INSERT INTO SubjectTeacher (subj_id,tch_id) VALUES ("+sid+","+tmpTID+")");
            else
              execQuery("UPDATE SubjectTeacher SET SubjectTeacher.tch_id = "+
                      tmpTID+" WHERE (((SubjectTeacher.subj_id)="+sid+"))");
          }

          break;
        }// </editor-fold>
        // <editor-fold defaultstate="collapsed" desc="4-room">
        case 4: {
          String tmpRoom=editValue;

          if(tmpRoom.equals("1"))
            tmpRoom="111(1)";
          if(tmpRoom.equals("2"))
            tmpRoom="222(1)";
          if(tmpRoom.equals("3"))
            tmpRoom="333(1)";
          cb.getEditor().setItem(tmpRoom);

          StringBuffer s=new StringBuffer(tmpRoom);
          boolean wrongRoom=false;

          for(int i=0;i<s.length();i++){
            char ch=s.charAt(i);
            if(!((ch>='0'&ch<='9')||(ch=='('|ch==')')||(ch>='a'&ch<='z')||(ch=='a'))){
              wrongRoom=true;
              break;
            }
          }

          if(s.length()<3||s.length()>7)
            wrongRoom=true;

          if(wrongRoom){
            cb.getEditor().selectAll();
            JOptionPane.showMessageDialog(mf," Correct room name has to contain: "+
                    "\n 2 or 3 digits for the room number "+
                    "\n and 1 digit for the building number, "+
                    "\n either in ( ) or without it"+
                    "\n\n\t examples: 11(1),123(2),233,2341","Wrong room name",JOptionPane.ERROR_MESSAGE);
            selectNextCell=false;
          }
          else{
            if(s.charAt(s.length()-1)!=')'){
              s.insert(s.length()-1,'(');
              s.append(')');
              cb.getEditor().setItem(s.toString());
            }

            s.deleteCharAt(s.length()-1);
            char par=s.charAt(s.indexOf("(")-1);
            if(par>'0'&par<'9')
              s.setCharAt(s.indexOf("("),' ');
            else
              s.deleteCharAt(s.indexOf("("));
            tmpRoom=s.toString();

            rs=execQuery("SELECT [3Rooms].name FROM 3Rooms WHERE ((([3Rooms].name)='"+tmpRoom+"'))");
            String tmpRoomCapacity;
            if(!rs.next()){
              Object val=tb.getValueAt(r,5);
              if(val!=null){
                if(val.toString().length()!=0)
                  tmpRoomCapacity=(String)tb.getValueAt(r,5);
                else
                  tmpRoomCapacity="0";
              }
              else
                tmpRoomCapacity="0";
              execQuery("INSERT INTO 3Rooms (name,capacity) VALUES ('"+tmpRoom+"',"+tmpRoomCapacity+")");

              rs=execQuery("SELECT SubjectRoom.subj_id FROM SubjectRoom "+
                      "WHERE (((SubjectRoom.subj_id)="+sid+"))");
              if(!rs.next())
                execQuery("INSERT INTO SubjectRoom (subj_id,aud_name) VALUES ("+sid+",'"+tmpRoom+"')");
              else
                execQuery("UPDATE SubjectRoom SET SubjectRoom.aud_name = '"+tmpRoom+"'"+
                        " WHERE (((SubjectRoom.subj_id)="+sid+"))");
            }
            else{
              rs=execQuery("SELECT SubjectRoom.subj_id FROM SubjectRoom "+
                      "WHERE (((SubjectRoom.subj_id)="+sid+"))");
              if(!rs.next()){
                execQuery("INSERT INTO SubjectRoom (subj_id,aud_name) VALUES ("+sid+",'"+tmpRoom+"')");
                c++;
              }
              else
                execQuery("UPDATE SubjectRoom SET SubjectRoom.aud_name = '"+tmpRoom+"'"+
                        " WHERE (((SubjectRoom.subj_id)="+sid+"))");

              rs=execQuery("SELECT [3Rooms].capacity FROM 3Rooms WHERE (([3Rooms].name)='"+tmpRoom+"')");
              rs.next();
              tb.setValueAt(rs.getString(1),r,c);
            }
          }

          break;
        }// </editor-fold>
        // <editor-fold defaultstate="collapsed" desc="5-room-capacity">
        case 5: {
//          System.out.println("cap");

          String tmpRoomCapacity=editValue;
          boolean wrongRoomCapacity=false;

          for(int i=0;i<tmpRoomCapacity.length();i++)
            if(!(tmpRoomCapacity.charAt(i)>='0'&&tmpRoomCapacity.charAt(i)<='9')){
              wrongRoomCapacity=true;
              break;
            }

          if(wrongRoomCapacity){
            tf.selectAll();
            JOptionPane.showMessageDialog(mf,"Enter a number","Wrong type",JOptionPane.ERROR_MESSAGE);
            selectNextCell=false;
          }
          else{
            String tmpRoom=(String)tb.getValueAt(r,4);
            StringBuffer s=new StringBuffer(tmpRoom);

            s.deleteCharAt(s.length()-1);
            char par=s.charAt(s.indexOf("(")-1);
            if(par>'0'&par<'9')
              s.setCharAt(s.indexOf("("),' ');
            else
              s.deleteCharAt(s.indexOf("("));
            tmpRoom=s.toString();

            execQuery("UPDATE 3Rooms SET [3Rooms].capacity = "+tmpRoomCapacity+" WHERE ((([3Rooms].name)='"+tmpRoom+"'))");
          }
          break;
        }// </editor-fold>
        // <editor-fold defaultstate="collapsed" desc="6-group">
        case 6: {
          String tmpGroup=editValue;
          String tmpGID;

          rs=execQuery("SELECT [4Groups].id FROM 4Groups WHERE ((([4Groups].name)='"+tmpGroup+"'))");
          if(!rs.next()){
            rs=execQuery("SELECT Max([4Groups].id) FROM 4Groups");
            rs.next();
            int intGID=rs.getInt(1);
            if(intGID==0)
              tmpGID="401";
            else
              tmpGID=String.valueOf(intGID+1);

            String tmpGroupCap;
            do{
              tmpGroupCap=(String)JOptionPane.showInputDialog(mf,"Enter the group capacity","New capacity",
                      JOptionPane.QUESTION_MESSAGE,null,null,25);
              for(int i=0;i<tmpGroupCap.length();i++)
                if(!(tmpGroupCap.charAt(i)>='0'&&tmpGroupCap.charAt(i)<='9')){
                  tmpGroupCap="";
                  break;
                }
            }while(tmpGroupCap.length()==0);
            execQuery("INSERT INTO 4Groups (id,name,capacity) VALUES ("+tmpGID+",'"+tmpGroup+"',"+tmpGroupCap+")");

            rs=execQuery("SELECT SubjectGroup.subj_id FROM SubjectGroup WHERE (((SubjectGroup.subj_id)="+sid+"))");
            if(!rs.next())
              execQuery("INSERT INTO SubjectGroup (subj_id,group_id) VALUES ("+sid+","+tmpGID+")");
            else{
              execQuery("UPDATE SubjectGroup SET SubjectGroup.group_id = "+tmpGID+" WHERE (((SubjectGroup.subj_id)="+sid+"))");

              boolean deleteGroup=false;
              if(tb.getRowCount()==1)
                deleteGroup=true;
              else if(rowNotComplete(r+1)){
                deleteGroup=true;
                r++;
              }

              if(deleteGroup){
                String delGroup=(String)lsGroups.getSelectedValue();
                execQuery("DELETE FROM 4Groups WHERE name='"+delGroup+"'");
              }
            }

            if(tb.getRowCount()==1)
              r++;

            newGroupAdded=true;
            getGroupNames();
            add();
          }
          else{
            tmpGID=rs.getString(1);
            execQuery("SELECT SubjectGroup.subj_id FROM SubjectGroup WHERE (((SubjectGroup.subj_id)="+sid+"))");
            if(!rs.next())
              execQuery("INSERT INTO SubjectGroup (subj_id,group_id) VALUES ("+sid+","+tmpGID+")");
            else
              execQuery("UPDATE SubjectGroup SET SubjectGroup.group_id = "+tmpGID+" WHERE (((SubjectGroup.subj_id)="+sid+"))");

            refresh();
            add();
            r=tb.getRowCount()-1;
          }

          break;
        }// </editor-fold>
      }

      if(!rowIsFull&selectNextCell){
        if(c==tb.getColumnCount()-1)
          c=0;
        c++;
      }
    }
    else{
//      System.out.println("empty");
    }

    if(!rowIsFull){
      if(!isTextField)
        cb.removeItemListener(itemListener);

      if(tb.editCellAt(r,c)){
        tb.changeSelection(r,c,false,false);
        tb.getEditorComponent().requestFocus();

        if(isTextField)
          tf.selectAll();
        else
          cb.getEditor().selectAll();
      }
    }
  }

  ResultSet execQuery(String query,boolean ... scroll){
    try{
      ResultSet rs=null;
      Statement st=null;

      if(con==null)
        return null;

      if(scroll.length==0)
        st=con.createStatement();
      else
        st=con.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_READ_ONLY);
      boolean ret=st.execute(query);
      if(ret)
        rs=st.getResultSet();

      return rs;
    }catch(SQLException e){
      e.printStackTrace();
      return null;
    }
  }

  boolean rowNotComplete(int row){
    for(int i=1;i<tb.getColumnCount();i++){
      Object val=tb.getValueAt(row,i);
      if(val==null)
        return true;
      else if(val.toString().length()==0)
        return true;
    }
    return false;
  }
}
