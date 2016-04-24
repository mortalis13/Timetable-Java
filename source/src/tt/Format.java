package tt;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

public class Format {
  final short BHAIR=CellStyle.BORDER_HAIR;
  final short BTHIN=CellStyle.BORDER_THIN;
  final short BTHICK=CellStyle.BORDER_THICK;
  final int LESSONS_ADAY=Timetable.LESSONS_ADAY;
  final int TOTAL_ROWS=5*2*LESSONS_ADAY+1;
  final String[] DAYS={"MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY"};

  MainForm mf=MainForm.mf;
  HSSFWorkbook w1;
  HSSFSheet s1;
  
  HSSFCellStyle flStyle;                                                                                               // full lessons
  HSSFCellStyle oetlStyle;                                                                                           // odd-even top lessons
  HSSFCellStyle oeblStyle;                                                                                           // odd-even bottom lessons
  HSSFCellStyle oehgtlStyle;                                                                                       // odd-even half-group top lessons
  HSSFCellStyle oehgblStyle;                                                                                       // odd-even half-group bottom lessons
  HSSFCellStyle daynumStyle;                                                                                       // days and lessons numbers
  HSSFCellStyle rightBorderStyle;                                                                              // right border of the group timetable (thick)
  HSSFCellStyle bottomBorderStyle;                                                                            // border of each lesson (thin)
  HSSFCellStyle lastLessonBorderStyle;                                                                     // last lesson of each day border (thick)

 // static{

  void initStyles(){
    w1=Connections.w1;
    s1=Connections.s1;

    int flFont=14;
    int oeFont=12;
    int oehgFont=10;

    flStyle=setStyle(flFont,"center","center","wrap");
    oetlStyle=setStyle(oeFont,"center","center","wrap");
    oetlStyle.setBorderBottom(BHAIR);                                                                                   // the line of division between two 18h lessons

    oeblStyle=setStyle(oeFont,"center","center","wrap");
    oeblStyle.setBorderTop(BHAIR);                                                                                   // the line of division between two 18h lessons

    oehgtlStyle=setStyle(oehgFont,"center","center","wrap");
    oehgtlStyle.setBorderRight(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgtlStyle.setBorderBottom(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgtlStyle.setBorderLeft(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgtlStyle.setBorderTop(BHAIR);                                                                                   // the line of division between two 18h lessons

    oehgblStyle=setStyle(oehgFont,"center","center","wrap");
    oehgblStyle.setBorderRight(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgblStyle.setBorderBottom(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgblStyle.setBorderLeft(BHAIR);                                                                                   // the line of division between two 18h lessons
    oehgblStyle.setBorderTop(BHAIR);                                                                                   // the line of division between two 18h lessons

    daynumStyle=w1.createCellStyle();
    daynumStyle.setBorderLeft(BTHICK);
    daynumStyle.setBorderRight(BTHICK);
    daynumStyle.setBorderTop(BTHICK);
    daynumStyle.setBorderBottom(BTHICK);

    rightBorderStyle=w1.createCellStyle();
    rightBorderStyle.setBorderRight(BTHICK);

    bottomBorderStyle=w1.createCellStyle();
    bottomBorderStyle.setBorderBottom(BTHIN);

    lastLessonBorderStyle=w1.createCellStyle();
    lastLessonBorderStyle.setBorderBottom(BTHICK);
    lastLessonBorderStyle.setBorderRight(BTHICK);
    lastLessonBorderStyle.setBorderLeft(BTHICK);
  }

  void formatFL(int row,int col,HSSFSheet s1){                                                         // Full Lessons (36h)
    s1.addMergedRegion(new CellRangeAddress(row,row+1,col,col+1));
    s1.getRow(row).getCell(col).setCellStyle(flStyle);
    s1.getRow(row).createCell(col+1).setCellStyle(flStyle);
    s1.getRow(row+1).createCell(col+1).setCellStyle(flStyle);
  }

  void formatOEL(int row,int col,int lesType,HSSFSheet s1){                                  // Odd-Even Lessons (18h)
    switch(lesType){
      case 1:                                                                                                                                  // (1 - the upper half; 2 - the lowe half)
        s1.addMergedRegion(new CellRangeAddress(row,row,col,col+1));
        s1.addMergedRegion(new CellRangeAddress(row+1,row+1,col,col+1));
        s1.getRow(row).getCell(col).setCellStyle(oetlStyle);
        s1.getRow(row).createCell(col+1).setCellStyle(oetlStyle);
        break;
      case 2:
        s1.addMergedRegion(new CellRangeAddress(row,row,col,col+1));
        s1.getRow(row).getCell(col).setCellStyle(oeblStyle);
        s1.getRow(row).createCell(col+1).setCellStyle(oeblStyle);
        break;
    }
  }

  void formatHGL(int row,int col,int load,int oelType,int lesSide,HSSFSheet s1){   // Half-Group Lessons (occupy 2 rows in one excel column)
    if(load==18){
      HSSFCellStyle cs=setStyle(10,"center","center","wrap");
      s1.getRow(row).getCell(col).setCellStyle(cs);

      int col1=col;
      switch(lesSide){
      case 0:                                                                                                                                  // PairType indicates the upper or lowe half of a pair-cell
        s1.getRow(row).getCell(col).setCellStyle(oehgtlStyle);
        s1.getRow(row).createCell(col+1).setCellStyle(w1.createCellStyle());
        setBorders(row+1,col+1,BHAIR);
        col1=col;
        break;
      case 1:
        s1.getRow(row).getCell(col).setCellStyle(oehgblStyle);
        setBorders(row+1,col-1,BHAIR);
        col1=col-1;
        break;
      }

      int id=-1;
      int x=s1.getNumMergedRegions();
      for(int i=0;i<x;i++){
        CellRangeAddress cra=s1.getMergedRegion(i);
        if(cra.getFirstRow()==row && cra.getFirstColumn()==col1)
          id=i;
      }
      if(id>-1)
        s1.removeMergedRegion(id);
    }
    else{
      HSSFCellStyle cs=setStyle(12,"center","center","wrap");
      switch(lesSide){
        case 0:
          s1.addMergedRegion(new CellRangeAddress(row,row+1,col,col));
          s1.addMergedRegion(new CellRangeAddress(row,row+1,col+1,col+1));
          cs.setBorderRight(BHAIR);
          break;
        case 1:
          s1.addMergedRegion(new CellRangeAddress(row,row+1,col-1,col-1));
          s1.addMergedRegion(new CellRangeAddress(row,row+1,col,col));
          cs.setBorderLeft(BHAIR);
          break;
      }
      s1.getRow(row).getCell(col).setCellStyle(cs);
      s1.getRow(row+1).createCell(col).setCellStyle(cs);
    }
  }

  void formatSheet(){
    int dayFont=24;                                                                                                                      // days
    int numFont=20;                                                                                                                      // lesson numbers
    int rowHeight=43;                                                                                                                  // each row
    int colWidth=28*200;                                                                                                             // lessons columns
    int fsColsWidth=28*50;                                                                                                         // first/second columns widths
    int groupColumn=2;                                                                                                                 // 0-based group column index

      // Days
    HSSFCellStyle cs=setStyle(dayFont,"center","center","bold","vorient");
    for(int i=0;i<5;i++){
      int dayRow=1+i*2*LESSONS_ADAY;                                                                                        // start from row 1 (0-based) end get rows 1,11,21,31,...
      HSSFRow r=s1.getRow(dayRow);
      HSSFCell c=r.createCell(0);
      c.setCellValue(DAYS[i]);
      c.setCellStyle(cs);
      s1.addMergedRegion(new CellRangeAddress(dayRow,1+(i+1)*2*LESSONS_ADAY-1,0,0));
    }

      // Lesson numbers
    cs=setStyle(numFont,"center","center","bold");
    for(int i=0;i<5;i++)
      for(int j=0;j<LESSONS_ADAY;j++){
        int dayRow=(j*2+1)+i*2*LESSONS_ADAY;                                                                            // rows 1,3,5,...; (lessons counting)+(day basis)
        HSSFRow r=s1.getRow(dayRow);
        HSSFCell c=r.createCell(1);
        c.setCellValue(j+1);
        c.setCellStyle(cs);
        s1.addMergedRegion(new CellRangeAddress(dayRow,(j*2+2)+i*2*LESSONS_ADAY,1,1));
      }

      // Rows Height; Days, Lesson numbers borders
    for(int i=0;i<TOTAL_ROWS;i++){
      HSSFRow r=s1.getRow(i);
      r.setHeightInPoints(rowHeight);
      if(i==0)
        r.setHeightInPoints(25);
      setBorders(i,0,BTHICK);                                                                                                                 // days
      setBorders(i,1,BTHICK);                                                                                                                 // lessons numbers
    }

      // Group names, Format empty cells, Draw frames around days and whole TTs
    for(int gr=0;gr<Connections.selectedGroupsCount;gr++){
      s1.addMergedRegion(new CellRangeAddress(0,0,groupColumn,groupColumn+1));
      s1.setColumnWidth(groupColumn,colWidth);                                                                     // first & second group columns widths
      s1.setColumnWidth(groupColumn+1,colWidth);

      HSSFCell c=s1.getRow(0).createCell(groupColumn);                                                       // group name
      c.setCellValue((String)mf.lsGroups.getModel().getElementAt(Connections.selectedGroupsIndexes[gr]));
      c.setCellStyle(cs);
      setBorders(0,groupColumn,BTHICK);                                                                                               // borders for two cells of the group name
      setBorders(0,groupColumn+1,BTHICK);

      for(int j=1;j<TOTAL_ROWS;j++){                                                                                        // set the right border of a group timtable as thick
        HSSFRow r=s1.getRow(j);
        c=r.getCell(groupColumn+1);
        if(c!=null)
          c.getCellStyle().setBorderRight(BTHICK);
        else
          r.createCell(groupColumn+1).setCellStyle(rightBorderStyle);
      }

      for(int j=2;j<TOTAL_ROWS;j+=2){                                                                                      // set the bottom border of each lesson as thin
        HSSFRow r=s1.getRow(j);
        HSSFCell c1=r.getCell(groupColumn);
        HSSFCell c2=r.getCell(groupColumn+1);
        if(c1!=null)
          c1.getCellStyle().setBorderBottom(BTHIN);
        else
          r.createCell(groupColumn).setCellStyle(bottomBorderStyle);
        c2.getCellStyle().setBorderBottom(BTHIN);
      }

      for(int j=10;j<TOTAL_ROWS;j+=10){                                                                                   // set borders of the last lessons on each day (the 5th lesson)
        HSSFRow r=s1.getRow(j);
        s1.addMergedRegion(new CellRangeAddress(j-1,j,groupColumn,groupColumn+1));
        r.createCell(groupColumn).setCellStyle(lastLessonBorderStyle);
        r.createCell(groupColumn+1).setCellStyle(lastLessonBorderStyle);
      }

      for(int j=Timetable.lastEmptyRow[gr];j<TOTAL_ROWS;j+=2)                                          // merge last empty cells on Friday
        s1.addMergedRegion(new CellRangeAddress(j,j+1,groupColumn,groupColumn+1));

      groupColumn+=2;
    }

    s1.setColumnWidth(0,fsColsWidth);  
    s1.setColumnWidth(1,fsColsWidth);
    s1.setZoom(85,100);
  }
  
  void setBorders(int row,int col,int border){
    HSSFRow r=s1.getRow(row);
    HSSFCell c=r.getCell(col);

    if(c!=null){
      HSSFCellStyle cs1=c.getCellStyle();
      cs1.setBorderLeft((short)border);
      cs1.setBorderRight((short)border);
      cs1.setBorderTop((short)border);
      cs1.setBorderBottom((short)border);
    }
    else{
      HSSFCellStyle cs1=w1.createCellStyle();
      cs1.setBorderLeft((short)border);
      cs1.setBorderRight((short)border);
      cs1.setBorderTop((short)border);
      cs1.setBorderBottom((short)border);
      r.createCell(col).setCellStyle(cs1);
    }
  }

  HSSFCellStyle setStyle(int fz,String align,String valign,String... other){
    HSSFCellStyle cs=w1.createCellStyle();
    if(fz!=0){
      HSSFFont f=w1.createFont();
      f.setFontHeightInPoints((short)fz);
      cs.setFont(f);
      for(String s:other){
        if(s.equals("bold"))
          f.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        if(s.equals("italic"))
          f.setItalic(true);
      }
    }

    if(align.equals("center"))
      cs.setAlignment(CellStyle.ALIGN_CENTER);
    if(valign.equals("center"))
      cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

    for(String s:other){
      if(s.equals("wrap"))
        cs.setWrapText(true);
      if(s.equals("vorient"))
        cs.setRotation((short)0xff);                                                                                         // vertical text orientation
    }
    return cs;
  }

  HSSFCellStyle formatStats(){
    HSSFCellStyle cs=w1.createCellStyle();
    HSSFFont f=w1.createFont();
    f.setBoldweight(Font.BOLDWEIGHT_BOLD);
    f.setFontHeightInPoints((short)12);
    cs.setFont(f);
    return cs;
  }
}
