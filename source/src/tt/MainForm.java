package tt;

import java.awt.Color;
import java.awt.Toolkit;
import java.awt.event.KeyEvent;

import javax.swing.ImageIcon;

public class MainForm extends javax.swing.JFrame {
  static MainForm mf;
  Connections con;
  Timetable tt;
  Format fmt;

  String title=" University Timetable";
  int pnum=0;

  String EditedValue, SelectedCellValue;
  int EditedRow, EditedCol, PrintedTTsNumber;

  public MainForm() {
    getContentPane().setBackground(new Color(250, 250, 250));
    initComponents();

    setLocation(70,80);
    setTitle(title);
    setIconImage(Toolkit.getDefaultToolkit().getImage("resources/icons/icon.png"));

    bNew.setIcon(new ImageIcon(Toolkit.getDefaultToolkit().getImage("resources/icons/1-new.png")));
    bLoad.setIcon(new ImageIcon(Toolkit.getDefaultToolkit().getImage("resources/icons/2-load.png")));
    bSave.setIcon(new ImageIcon(Toolkit.getDefaultToolkit().getImage("resources/icons/3-save.png")));
    bGenerate.setIcon(new ImageIcon(Toolkit.getDefaultToolkit().getImage("resources/icons/4-generate.png")));
    bExit.setIcon(new ImageIcon(Toolkit.getDefaultToolkit().getImage("resources/icons/5-exit.png")));
  }
  // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
  private void initComponents() {

    jScrollPane1 = new javax.swing.JScrollPane();
    lsGroups = new javax.swing.JList();
    jScrollPane2 = new javax.swing.JScrollPane();
    tbData = new javax.swing.JTable();
    bNew = new javax.swing.JButton();
    bExit = new javax.swing.JButton();
    bGenerate = new javax.swing.JButton();
    bSave = new javax.swing.JButton();
    bLoad = new javax.swing.JButton();
    jLabel1 = new javax.swing.JLabel();
    jLabel2 = new javax.swing.JLabel();
    bAdd = new javax.swing.JButton();
    bDelete = new javax.swing.JButton();
    bRefresh = new javax.swing.JButton();

    setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
    setTitle("University Timetable");
    addWindowListener(new java.awt.event.WindowAdapter() {
      public void windowClosing(java.awt.event.WindowEvent evt) {
        formWindowClosing(evt);
      }
      public void windowOpened(java.awt.event.WindowEvent evt) {
        formWindowOpened(evt);
      }
    });
    addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        formKeyPressed(evt);
      }
    });

    lsGroups.addListSelectionListener(new javax.swing.event.ListSelectionListener() {
      public void valueChanged(javax.swing.event.ListSelectionEvent evt) {
        lsGroupsValueChanged(evt);
      }
    });
    lsGroups.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        lsGroupsKeyPressed(evt);
      }
    });
    jScrollPane1.setViewportView(lsGroups);

    tbData.setModel(new javax.swing.table.DefaultTableModel(
      new Object [][] {
        {},
        {},
        {},
        {}
      },
      new String [] {

      }
    ));
    jScrollPane2.setViewportView(tbData);

    bNew.setIcon(new javax.swing.ImageIcon("D:\\Documentos\\Netbeans\\Timetable\\icons\\1-new.png")); // NOI18N
    bNew.setText("New");
    bNew.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
    bNew.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bNewActionPerformed(evt);
      }
    });
    bNew.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bNewKeyPressed(evt);
      }
    });

    bExit.setIcon(new javax.swing.ImageIcon("D:\\Documentos\\Netbeans\\Timetable\\icons\\5-exit.png")); // NOI18N
    bExit.setText("Exit");
    bExit.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
    bExit.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bExitActionPerformed(evt);
      }
    });
    bExit.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bExitKeyPressed(evt);
      }
    });

    bGenerate.setIcon(new javax.swing.ImageIcon("D:\\Documentos\\Netbeans\\Timetable\\icons\\4-generate.png")); // NOI18N
    bGenerate.setText("Generate");
    bGenerate.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
    bGenerate.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bGenerateActionPerformed(evt);
      }
    });
    bGenerate.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bGenerateKeyPressed(evt);
      }
    });

    bSave.setIcon(new javax.swing.ImageIcon("D:\\Documentos\\Netbeans\\Timetable\\icons\\3-save.png")); // NOI18N
    bSave.setText("Save");
    bSave.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
    bSave.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bSaveActionPerformed(evt);
      }
    });
    bSave.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bSaveKeyPressed(evt);
      }
    });

    bLoad.setIcon(new javax.swing.ImageIcon("D:\\Documentos\\Netbeans\\Timetable\\icons\\2-load.png")); // NOI18N
    bLoad.setText("Load");
    bLoad.setHorizontalAlignment(javax.swing.SwingConstants.LEFT);
    bLoad.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bLoadActionPerformed(evt);
      }
    });
    bLoad.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bLoadKeyPressed(evt);
      }
    });

    jLabel1.setText("Groups");

    bAdd.setText("Add");
    bAdd.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bAddActionPerformed(evt);
      }
    });
    bAdd.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bAddKeyPressed(evt);
      }
    });

    bDelete.setText("Delete");
    bDelete.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bDeleteActionPerformed(evt);
      }
    });
    bDelete.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bDeleteKeyPressed(evt);
      }
    });

    bRefresh.setText("Refresh");
    bRefresh.addActionListener(new java.awt.event.ActionListener() {
      public void actionPerformed(java.awt.event.ActionEvent evt) {
        bRefreshActionPerformed(evt);
      }
    });
    bRefresh.addKeyListener(new java.awt.event.KeyAdapter() {
      public void keyPressed(java.awt.event.KeyEvent evt) {
        bRefreshKeyPressed(evt);
      }
    });

    javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
    getContentPane().setLayout(layout);
    layout.setHorizontalGroup(
      layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
      .addGroup(layout.createSequentialGroup()
        .addContainerGap()
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.CENTER)
            .addComponent(bNew, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addComponent(bLoad, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addComponent(bExit, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE))
          .addComponent(jLabel2)
          .addComponent(bSave, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addComponent(jLabel1)
          .addComponent(bGenerate, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 115, javax.swing.GroupLayout.PREFERRED_SIZE))
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addGroup(layout.createSequentialGroup()
            .addComponent(bAdd)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(bDelete)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(bRefresh))
          .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 977, Short.MAX_VALUE))
        .addContainerGap())
    );

    layout.linkSize(javax.swing.SwingConstants.HORIZONTAL, new java.awt.Component[] {bExit, bLoad, bNew, bSave});

    layout.setVerticalGroup(
      layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
      .addGroup(layout.createSequentialGroup()
        .addContainerGap()
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 591, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addGroup(layout.createSequentialGroup()
            .addComponent(bNew, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addComponent(bLoad, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addComponent(bSave, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
            .addComponent(bGenerate, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addGap(13, 13, 13)
            .addComponent(bExit, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addGap(30, 30, 30)
            .addComponent(jLabel1)
            .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
            .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 101, javax.swing.GroupLayout.PREFERRED_SIZE)
            .addGap(76, 76, 76)
            .addComponent(jLabel2)))
        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
          .addComponent(bAdd, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addComponent(bDelete, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
          .addComponent(bRefresh, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE))
        .addContainerGap(20, Short.MAX_VALUE))
    );

    pack();
  }// </editor-fold>//GEN-END:initComponents

  private void formWindowOpened(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowOpened
    con=new Connections();
    tt=new Timetable();
    fmt=new Format();

    con.initTable();
    con.newTimetable();

//    bLoadActionPerformed(null);
  }//GEN-LAST:event_formWindowOpened
  
  private void bLoadActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bLoadActionPerformed
    con.dbLoad();
    con.getGroupNames();

//    lsGroups.setSelectionInterval(0,1);
  }//GEN-LAST:event_bLoadActionPerformed

  private void lsGroupsValueChanged(javax.swing.event.ListSelectionEvent evt) {//GEN-FIRST:event_lsGroupsValueChanged
    if(lsGroups.getSelectedIndex()!=-1){
      con.fillTable();

//      con.tb.editCellAt(0,1);
//      con.tb.getEditorComponent().requestFocus();

//      bAddActionPerformed(null);
//      bGenerateActionPerformed(null);
    }
  }//GEN-LAST:event_lsGroupsValueChanged

  private void bGenerateActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bGenerateActionPerformed
    if(lsGroups.getSelectedIndex()!=-1){
//      jLabel2.setText("G-"+String.valueOf(++pnum));

      con.importData();
      tt.generateTimetablesArray();
      tt.findOptimalTimetable();

      con.xlConnect();
      tt.fillExcelSheet();
      fmt.formatSheet();
      con.openExcel();

//      this.toFront();
//      formWindowClosed(null);
//      System.exit(0);
    }
  }//GEN-LAST:event_bGenerateActionPerformed
  
  private void formKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_formKeyPressed
    if(evt.getKeyCode()==KeyEvent.VK_ESCAPE){
//      formWindowClosing(null);
//      System.exit(0);
    }
  }//GEN-LAST:event_formKeyPressed
  
  private void bLoadKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bLoadKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_bLoadKeyPressed
  
  private void lsGroupsKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_lsGroupsKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_lsGroupsKeyPressed

  private void bGenerateKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bGenerateKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_bGenerateKeyPressed

  private void bNewActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bNewActionPerformed
    con.newTimetable();
  }//GEN-LAST:event_bNewActionPerformed

  private void bNewKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bNewKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_bNewKeyPressed

  private void bSaveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bSaveActionPerformed
    con.dbSave();
  }//GEN-LAST:event_bSaveActionPerformed

  private void bSaveKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bSaveKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_bSaveKeyPressed

  private void bExitActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bExitActionPerformed
      formWindowClosing(null);
      System.exit(0);
  }//GEN-LAST:event_bExitActionPerformed

  private void bExitKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bExitKeyPressed
//    formKeyPressed(evt);
  }//GEN-LAST:event_bExitKeyPressed
  
  private void formWindowClosing(java.awt.event.WindowEvent evt) {//GEN-FIRST:event_formWindowClosing
    con.dbDisconnect(true);
//    try{
//      Runtime.getRuntime().exec("taskkill /f /im EXCEL.exe");
//    }catch(IOException e){
//      e.printStackTrace();
//    }
  }//GEN-LAST:event_formWindowClosing

  private void bAddActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bAddActionPerformed
    con.add(); 
  }//GEN-LAST:event_bAddActionPerformed

  private void bAddKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bAddKeyPressed
    // TODO add your handling code here:
  }//GEN-LAST:event_bAddKeyPressed

  private void bDeleteActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bDeleteActionPerformed
    con.delete();
  }//GEN-LAST:event_bDeleteActionPerformed

  private void bDeleteKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bDeleteKeyPressed
    // TODO add your handling code here:
  }//GEN-LAST:event_bDeleteKeyPressed

  private void bRefreshActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_bRefreshActionPerformed
    con.refresh();
  }//GEN-LAST:event_bRefreshActionPerformed

  private void bRefreshKeyPressed(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_bRefreshKeyPressed
    // TODO add your handling code here:
  }//GEN-LAST:event_bRefreshKeyPressed
    
  public static void main(String args[]) {
    //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
//                if ("Windows Classic".equals(info.getName())) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(MainForm.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                mf=new MainForm();
                mf.setVisible(true);
            }
        });
  }
  
  // Variables declaration - do not modify//GEN-BEGIN:variables
  public javax.swing.JButton bAdd;
  public javax.swing.JButton bDelete;
  public javax.swing.JButton bExit;
  public javax.swing.JButton bGenerate;
  public javax.swing.JButton bLoad;
  public javax.swing.JButton bNew;
  public javax.swing.JButton bRefresh;
  public javax.swing.JButton bSave;
  public javax.swing.JLabel jLabel1;
  public javax.swing.JLabel jLabel2;
  public javax.swing.JScrollPane jScrollPane1;
  public javax.swing.JScrollPane jScrollPane2;
  public javax.swing.JList lsGroups;
  public javax.swing.JTable tbData;
  // End of variables declaration//GEN-END:variables
}
