package tt;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;

public class Timetable{
  static class TTBElement{
    int[] timetable;
    int breaks;
  }
  static class Lesson implements Cloneable{
    String subject;
    int load;
    String teacher;
    String room;
    int roomCapacity;

    public Lesson clone() throws CloneNotSupportedException{
      return (Lesson)super.clone();
    }
  }

  final int PRINT_STATS=0;
  
  final int USE_SPECIFIED_TIMETABLES=0;                                                                    // to print the groups timetables with specified lesson indexes (without random indexes)
  final int ST_GROUP1=0;
  final int ST_GROUP2=1;
  final int[] SPECIFIED_TT1={14,3,19,5,0,13,21,20,17,10,16,4,2,18,9,6,11,15,12,8,1,7};
  final int[] SPECIFIED_TT2={7,0,5,12,18,17,11,10,19,1,6,14,13,15,3,16,9,2,4,8};
   
  static final int PRINT_TIMETABLE_COUNT=1;                                                                         // how many TTs we're to print out
  static final int SUBJECTS=30;                                                                                               // number of subjects
  static final int LESSONS_ADAY=5;                                                                                         // Subjects per 1 day
  static final int GROUPS=2;                                                                                                    // max Groups in the TimeTable
  static final int POPULATION=100;                                                                                         // TimeTables Population
  static final int RCONFLICT=1;                                                                                               // consts for detectConflicts function
  static final int TCONFLICT=2;

  static int[] lesCount=new int[GROUPS];                                                                                // lessons count
  static int[] groupsCapacity=new int[GROUPS];
  static int[] lastEmptyRow=new int[GROUPS];                                                                        // to format last empty cells on Friday
  int[] optimalTimetableIndexes=new int[GROUPS];                                                    // indexes in the 'ttb' array

  static Lesson[][] lessons=new Lesson[GROUPS][SUBJECTS];                                                // main array for (subject,teacher,room,...) elements
  TTBElement[][] ttb=new TTBElement[GROUPS][POPULATION];                                     // ttb - timetables (contain indexes of lessons from the 'lessons' array) and breaks for each timetable
  int[][][] dlp=new int[GROUPS][3][SUBJECTS];                                                         // dlp - day-lesson-lesson_part to detect breaks and fill an excel sheet
  int[][] conflictsDLP=new int[SUBJECTS][3];

  MainForm mf=MainForm.mf;

  void generateTimetablesArray(){                                                                              // generate random indexes of lessons from the 'lessons' array
    if(USE_SPECIFIED_TIMETABLES==1){
      ttb[ST_GROUP1][0]=new TTBElement();
      ttb[ST_GROUP1][0].timetable=SPECIFIED_TT1;
      optimalTimetableIndexes[ST_GROUP1]=0;

      ttb[ST_GROUP2][1]=new TTBElement();
      ttb[ST_GROUP2][1].timetable=SPECIFIED_TT2;
      optimalTimetableIndexes[ST_GROUP2]=1;
    }
    else{
      int[] includedLessonsIndexes=new int[SUBJECTS];                                                           // and fill the 'ttb' array with these indexes
      for(int grn=0;grn<Connections.selectedGroupsCount;grn++){
        int gr=Connections.selectedGroupsIndexes[grn];

        for(int i=0;i<POPULATION;i++){
          ttb[gr][i]=new TTBElement();
          ttb[gr][i].timetable=new int[SUBJECTS];

          for(int j=0;j<SUBJECTS;j++)
            includedLessonsIndexes[j]=0;                                                                                        // this array is used to aviod the including of the same indexes
          for(int j=0,x;j<lesCount[gr];j++){
            do{
              x=(int)(Math.random()*lesCount[gr]);
            }while(includedLessonsIndexes[x]==1);
            ttb[gr][i].timetable[j]=x;
            includedLessonsIndexes[x]=1;
          }
        }
      }
    }
  }

  void findOptimalTimetable(){
    if(USE_SPECIFIED_TIMETABLES==0){
      int minBreaks;
      int maxmins=10;
      int[] mbtc=new int[GROUPS];                                                                                                 // mbtc - min breaks tts count (number of timetables with min number of breaks for the each group)
      int[][] mbti=new int[GROUPS][POPULATION];                                                                       // mbti - min breaks tts indexes (indexes of timetables with min breaks)
      int[] mutationsCount=new int[GROUPS];

      for(int grn=0;grn<Connections.selectedGroupsCount;grn++){                                          // loop for the min breaks detecting
        int gr=Connections.selectedGroupsIndexes[grn];
        for(int i=0;i<POPULATION;i++)
          ttb[gr][i].breaks=detectBreaks(ttb[gr][i].timetable,gr);                                      // breaks count for each timetable, generated in the 'generateTimetablesArray()' method

        do{                                                                                                                                         // loop until min breaks count is equal to 0
          for(int i=0;i<POPULATION;i++)
            mbti[gr][i]=0;
          mbtc[gr]=1;
          mbti[gr][0]=0;
          minBreaks=ttb[gr][0].breaks;

          for(int i=1;i<POPULATION;i++){                                                                                      // find the minimum within generated timetables in the 'ttb' array
            if(ttb[gr][i].breaks<minBreaks){
              for(int j=0;j<mbtc[gr];j++)                                                                                       // clear found indexes if there are timetables with less breaks count
                mbti[gr][j]=0;
              mbti[gr][0]=i;
              mbtc[gr]=1;
              minBreaks=ttb[gr][i].breaks;
            }
            else if(ttb[gr][i].breaks==minBreaks&mbtc[gr]<=maxmins){                                                                      // store other timetables with the same number of breaks to choose a random timetable from them
              mbti[gr][mbtc[gr]]=i;
              mbtc[gr]++;
            }
          }

          if(minBreaks!=0){                                                                                                              // mutations
            mutationsCount[gr]++;
            for(int i=0;i<mbtc[gr];i++){
              mutation(ttb[gr][mbti[gr][i]].timetable,gr);
              ttb[gr][mbti[gr][i]].breaks=detectBreaks(ttb[gr][mbti[gr][i]].timetable,gr);
            }
          }
        }while(minBreaks>0);
        optimalTimetableIndexes[gr]=mbti[gr][(int)(Math.random()*mbtc[gr])];                  // these are indexes from 'ttb' with optimal timetables which will be printed
      }

      if(Connections.selectedGroupsCount>1){                                                                           // find conflicts netween two groups lessons
        int agtc=0;                                                                                                                           // agtc - all groups timetables count (number of TTs for all (2) groups. Is used to find TTs with min Conflicts)
        int minConflictsSum;
        int[][] ticc=new int[4][POPULATION*20];                                                                        // ticc - tts indexes and conflicts count (used to minimize conflicts between groups timetables)

        for(int i=0;i<mbtc[0];i++){
          fillDLP(ttb[0][mbti[0][i]].timetable,0);                                                                   // fill 'dlp' to compare lessons position in the first timetable and the second one
          for(int j=0;j<mbtc[1];j++){                                                                                            // compare each tt of the first group with each one of the second group
            fillDLP(ttb[1][mbti[1][j]].timetable,1);                                                                 //

            ticc[0][agtc]=mbti[0][i];                                                                                             // 0 - first tt index in the comparison
            ticc[1][agtc]=mbti[1][j];                                                                                             // 1 - seond tt index
            ticc[2][agtc]=detectConflicts(RCONFLICT,mbti[0][i],mbti[1][j],false);            // 2 - number of room conflicts between two tts
            ticc[3][agtc]=detectConflicts(TCONFLICT,mbti[0][i],mbti[1][j],false);            // 3 - number of teacher conflicts
            agtc++;
          }
        }

        minConflictsSum=ticc[2][0]+ticc[3][0];                                                                         // the sum of A-T conflicts is used
        optimalTimetableIndexes[0]=ticc[0][0];
        optimalTimetableIndexes[1]=ticc[1][0];

        for(int i=1;i<agtc;i++)                                                                                                     // find the minimum of the conflicts sums
          if(ticc[2][i]+ticc[3][i]<minConflictsSum){
            minConflictsSum=ticc[2][i]+ticc[3][i];
            optimalTimetableIndexes[0]=ticc[0][i];                                                                     // final optimal tt indexes
            optimalTimetableIndexes[1]=ticc[1][i];
          }
      }
    }
  }

  int detectBreaks(int[] timetableChromosome,int gr){                                            // Chromosome Fitness function - returns number of breaks for a given timetable
    int breaks=0;                                                                                                                          // breaks count
    int row=1;
    int counter=0;                                                                                                                        // global counter
    int les=0;                                                                                                                               // lesson index
    boolean firstOEL=false;                                                                                                       // first odd-even lesson was printed
    boolean firstHGL=false;                                                                                                       // first half group

    while(counter<lesCount[gr]){                                                                                             // the loop for all lessons for the current group
      if((row+1)%10!=0){                                                                                                             // if any 18h or 36h äóûûùò follows a HGL
        if(firstHGL==true)                                                                                                           // and now we are on the 3rd or 4th pair
          if(((row+5)%10==0)||((row+3)%10==0))                                                                        // then it is a break
            breaks++;

        int x=timetableChromosome[les];                                                                                    // next lesson index (in the lessons array)
        if((lessons[gr][x].load==36)||(lessons[gr][x].load==54)){                                    // 36h, 54h
          if(lessons[gr][x].roomCapacity>=(groupsCapacity[gr]-5)){                                  // full les
            if(firstOEL==true)                                                                                                        // top half 18h lesson was printed on the previous step
              if(((row-1)%10!=0)&&((row+7)%10!=0)){                                                                   // we are on the 1st or 2nd les
                firstOEL=false;
                breaks++;                                                                                                                    // a new break
              }
              else
                firstOEL=false;
            firstHGL=false;
            row=row+2;                                                                                                                      // was there any break or not - we're moving to the next les number
          }
          else{                                                                                                                                 // a half group 36h les
            if(firstOEL==true)                                                                                                       // prev les was 18h
              if((((row-1)%10!=0)&&(row+7)%10!=0)){                                                                  // 1st or 2nd les is now
                firstOEL=false;
                breaks++;                                                                                                                   // a new break
              }
              else
                firstOEL=false;
            firstHGL=true;                                                                                                               // we're taking into consideration that one of the HGPs was printed
            row=row+2;                                                                                                                      // and later we'll check if it makes a break relative to the next les
          }
        }
        else if(lessons[gr][x].roomCapacity>=(groupsCapacity[gr]-5)){                            // 18h non-HGL
          if(firstOEL==false){                                                                                                      // it's the 1st 18h les in the current cell
            firstOEL=true;
            row=row+2;                                                                                                                      // next cell
          }
          else
            firstOEL=false;                                                                                                             // if there is already the top 18h les in this cell
          firstHGL=false;                                                                                                               // then a second one (bottom les) is going to be printed now
        }
        else{                                                                                                                                    // HGL 18h les
          if(firstHGL==false){
            firstHGL=true;                                                                                                               // we've just printed 1 HGL
            if(firstOEL==false){                                                                                                    // working with HGL 18h lessons
              firstOEL=true;
              row+=2;
            }
            else
              firstOEL=false;
          }
          else{                                                                                                                                  // an HGL was printed previously
            if(firstOEL==false){
              firstOEL=true;                                                                                                             // in this branch prev HGP was printed in the prev pair-cell
              row+=2;                                                                                                                          // and we leave the 'firstHGL' value in true
              }
            else{
              firstHGL=false;                                                                                                           // here the second HGL is printed in the same cell
              firstOEL=false;                                                                                                           // as the first one
              }                                                                                                                                    // and we are to set 'firstHGL' flag to false
            }                                                                                                                                      // 'cause in the other case it could be a double break in the TimeTable here
          }
        counter++;
        les++;                                                                                                                                 // to the next les index
      }
      else
        row=row+2;                                                                                                                         // skip the 5th les
    }
    return breaks;
  }

  int detectConflicts(int type,int tt1,int tt2,boolean fillConflictsDLP){        // Detect number of conflicts between two groups TTS
    int conflicts=0;                                                                         // conflicts count     // ConflictType:
    int gr1=0;                                                                                    // group indexes       //              1-for similar Rooms
    int gr2=1;                                                                                                                              //              2-for similar Teachers
    boolean compared;

    if(fillConflictsDLP)
      for(int i=0;i<SUBJECTS;i++)
        for(int j=0;j<3;j++)
          conflictsDLP[i][j]=0;

    for(int les1=0;les1<lesCount[gr1];les1++){                                                                    // the 1st group index
      compared=false;
      int x1=ttb[gr1][tt1].timetable[les1];                                                                           // les index in the 'lessons' array for the 1st group

      for(int les2=0;les2<lesCount[gr2];les2++)                                                                    // the 2nd group index
        if((dlp[gr1][0][les1]==dlp[gr2][0][les2])&&
              (dlp[gr1][1][les1]==dlp[gr2][1][les2])&&
              (dlp[gr1][2][les1]+dlp[gr2][2][les2]!=3)){    // if days and lessons numbers for current cell coincide

          int x2=ttb[gr2][tt2].timetable[les2];                                                                                                                                                                    // and if there are 18h lessons which satisfy this condition,
          switch(type){                                                                                                                                                                                                               // then these 18h lessons aren't in different parts of the same cell
            case RCONFLICT:{                                                                                                                                                                                                       // (top-bottom or bottom-top) but only in the same row
              if((lessons[gr1][x1].room.equals(lessons[gr2][x2].room))
                      &&(!lessons[gr1][x1].room.equals(""))){                                          // if rooms coincide
                if(fillConflictsDLP==true)
                  for(int i=0;i<3;i++)
                    conflictsDLP[conflicts][i]=dlp[gr1][i][les1];
                conflicts++;                                                                                                                                                                                                      // there is a conflict
              }
              break;
            }
            case TCONFLICT:{
              if((lessons[gr1][x1].subject.equals("Physical Training"))&&(lessons[gr2][x2].subject.equals("Physical Training"))){
                String teacher1=lessons[gr1][x1].teacher;                                                          // PT lessons have 2 teachers and we have to compare all of them between 2 groups
                String teacher2=lessons[gr2][x2].teacher;

                String s1=teacher1;
                String s2=teacher2;

                String tch11=s1.substring(0,s1.indexOf(','));
                String tch12=s1.substring(s1.indexOf(',')+2);
                String tch21=s2.substring(0,s2.indexOf(','));
                String tch22=s2.substring(s2.indexOf(','));

                if((tch11.equals(tch21))||(tch11.equals(tch22))||(tch12.equals(tch21))||(tch12.equals(tch22))){
                  if(fillConflictsDLP==true)
                    for(int i=0;i<3;i++)
                      conflictsDLP[conflicts][i]=dlp[gr1][i][les1];
                  conflicts++;
                }
              }
              else if(lessons[gr1][x1].teacher.equals(lessons[gr2][x2].teacher)){
                if(fillConflictsDLP==true)
                  for(int i=0;i<3;i++)
                    conflictsDLP[conflicts][i]=dlp[gr1][i][les1];
                conflicts++;
              }
              break;
            }
          }
          compared=true;                                                                                                                // we have compared lessons in the current row
        }
        else if(compared==true)                                                                                                    // if one pair of lessons was compared
          break;                                                                                                                               // and there are no lessons with the same day-les_num
    }                                                                                                                                              // then we can interrupt current 'for les2' cycle
    return conflicts;                                                                                                                  // ('cause there won't be any coincidence later in the second group array)
  }

  void mutation(int[] timetableChromosome,int gr){                                                 // Mutation - xchanging 2 random pairs indexes
    int x, y;
    x=(int)(Math.random()*lesCount[gr]);
    do{
      y=(int)(Math.random()*lesCount[gr]);
    }while(x==y);
    int tmp=timetableChromosome[x];
    timetableChromosome[x]=timetableChromosome[y];
    timetableChromosome[y]=tmp;
  }

// *********** fillDLP() - fills DPP 3d array which is used later for an excel sheet filling and to find conflicts ************
// DLP:
// 0 - group index (gr)
// 1 - 3 columns:
//       1-day number (day)
//       2-lesson number on the current day (lesn)
//       3-lesson part (if it 36h = 0, 18h =
//                                          (1-for the top half)
//                                          (2-for the bottom half)
// 2 - lesson index for current group (les)
  void fillDLP(int[] timetable,int gr){                                                                     
    int les=0;                                                                                                                               // lesson index (corresponds with the 'lessons' array)
    int day=1;                                                                                                                               // day of the week number
    int lesn=1;                                                                                                                             // lesson number (on the current day of the week)
    boolean filled=false;
    boolean firstOEL=false;

    for(int i=0;i<3;i++)
      for(int j=0;j<SUBJECTS;j++)
        dlp[gr][i][j]=0;

    while(les<lesCount[gr]){                                                                                                    // cycle for all lessons of the current group
      filled=false;                                                                                                                      // if current cell in an excel sheet is filled (1 36h lesson, 1 or 2 18h lessons)
      while(filled==false){
        int x=timetable[les];                                                                                                      // the next index from the array that points to a lesson in the 'lessons' array
        if((lessons[gr][x].load==36)||(lessons[gr][x].load==54)){   /*36-54h*/
          filled=true;                                                                                                                    // DLP column for current lesson is filled with 3 values and we can move to the next lesson in the TT (in an Excel sheet)
          if(firstOEL==true)                                                                                                         // if prev pair was 18 then
            firstOEL=false;                                                                                                            // we're to inc 'lesn' ('cause now 'lesn' points to the prev (wrong) pair)
          else{
            dlp[gr][0][les]=day;
            dlp[gr][1][les]=lesn;
            dlp[gr][2][les]=0;
            les++;                                                                                                                             // the next lesson index
          }
        }
        else{                                                                                                      /*18h*/
          if(firstOEL==false){                                                                                                      // first (top) 18h lesson in a lesson-cell
            dlp[gr][0][les]=day;
            dlp[gr][1][les]=lesn;
            dlp[gr][2][les]=1;

            firstOEL=true;                                                                                                               // lesson-cell in the Excel sheet is not filled yet
            if(les==(lesCount[gr]-1))
              filled=true;
          }
          else{
            dlp[gr][0][les]=day;
            dlp[gr][1][les]=lesn;
            dlp[gr][2][les]=2;

            firstOEL=false;
            filled=true;                                                                                                                  // we can move to the next excel lesson-cell
          }
          les++;
        }
      }

      lesn++;                                                                                                                                 // to the next lesson number in the excel sheet
      if(lesn==5){
        lesn=1;                                                                                                                               // only 4 lessons on each day are used
        day++;
      }
    }
  }
  
  void fillExcelSheet(){
    int groupColumn=2;                                                                                                              // initial group column in the Excel sheet (for the 1st group)

    HSSFSheet s1=Connections.s1;
    mf.fmt.initStyles();

    for(int i=0;i<51;i++)
      s1.createRow(i);

    for(int grn=0;grn<Connections.selectedGroupsCount;grn++){                                         // main printing cycle for TTs of selected groups
      int row=0;                                                                                                                             // Excel row number to print lessons
      boolean firstHGL=false;                                                                                                     // flag for HGLs

      int gr=Connections.selectedGroupsIndexes[grn];                                                           // next group index
      fillDLP(ttb[gr][optimalTimetableIndexes[gr]].timetable,gr);                                   // filling up the DÄP array with day-les-les_part values for each group TT
      for(int les=0;les<lesCount[gr];les++){                                                                          // cycle for lessons indexes from indexes array (TTB)
        String lessonString;
        int x=ttb[gr][optimalTimetableIndexes[gr]].timetable[les];                                   // next lesson index for 'lessons' array
        
        switch(dlp[gr][2][les]){                                                                                                 // the 3rd row in the DLP responses for load and les-part (top/bottom) (36h-18h[1]-18h[2])
          case 0:{                                                                                  // *36h*
            row=(dlp[gr][0][les]-1)*10+dlp[gr][1][les]*2-1;
            lessonString=lessons[gr][x].subject+' '+lessons[gr][x].teacher+' '+lessons[gr][x].room;

            if(lessons[gr][x].roomCapacity>=(groupsCapacity[gr]-5)){
              s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);        // printing out into the Excel sheet
              mf.fmt.formatFL(row,groupColumn,s1);
            }
            else{                                                                                                                               // 36h HGLs
              if(firstHGL==false){                                                                                                  // left-side HGL (first)
                s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn,36,0,0,s1);
                firstHGL=true;
              }
              else{                                                                                                                             // right-side HGLs
                s1.getRow(row).createCell(groupColumn+1).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn+1,36,0,1,s1);
                firstHGL=false;
              }
            }
            break;
          }
          case 1:{                                                                                   // *18h top*
            row=(dlp[gr][0][les]-1)*10+dlp[gr][1][les]*2-1;
            lessonString=lessons[gr][x].subject+" odd weeks "+lessons[gr][x].teacher+' '+lessons[gr][x].room;

            if(lessons[gr][x].roomCapacity>=(groupsCapacity[gr]-5)){
              s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);
              mf.fmt.formatOEL(row,groupColumn,1,s1);
            }
            else{                                                                                                                               // 18h top HGLs
              if(firstHGL==false){
                s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn,18,1,0,s1);
                firstHGL=true;
              }
              else{
                s1.getRow(row).createCell(groupColumn+1).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn+1,18,1,1,s1);
                firstHGL=false;
              }
            }
            break;
          }
          case 2:{                                                                                  // *18h bottom*
            row=(dlp[gr][0][les]-1)*10+dlp[gr][1][les]*2;
            lessonString=lessons[gr][x].subject+" even weeks "+lessons[gr][x].teacher+' '+lessons[gr][x].room;

            if(lessons[gr][x].roomCapacity>=(groupsCapacity[gr]-5)){
              s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);
              mf.fmt.formatOEL(row,groupColumn,2,s1);
            }
            else{                                                                                                                               // 18h bottom HGLs
              if(firstHGL==false){
                s1.getRow(row).createCell(groupColumn).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn,18,2,0,s1);
                firstHGL=true;
              }
              else{
                s1.getRow(row).createCell(groupColumn+1).setCellValue(lessonString);
                mf.fmt.formatHGL(row,groupColumn+1,18,2,1,s1);
                firstHGL=false;
              }
            }
            break;
          }
        }
      }
      groupColumn+=2;                                                                                                                   // to the next group column in the sheet (each group TT requires 2 Excel columns)
      if(row%2==0) lastEmptyRow[grn]=row+1;                                                                           // to format rest empty lesson-cells in the end of the week (in the 'mf.fmt.formatSheet' method)
      else lastEmptyRow[grn]=row+2;
    }

    if(PRINT_STATS==1){
      int x1=detectConflicts(RCONFLICT,optimalTimetableIndexes[0],optimalTimetableIndexes[1],true);
      s1.getRow(0).createCell(7).setCellValue("room: ");
      s1.getRow(0).createCell(8).setCellValue(x1);
      if(x1>0)
        for(int i=0;i<x1;i++)
          for(int j=0;j<3;j++)
            s1.getRow(i).createCell(10+j).setCellValue(conflictsDLP[i][j]);

      int x2=detectConflicts(TCONFLICT,optimalTimetableIndexes[0],optimalTimetableIndexes[1],true);
      s1.getRow(1).createCell(7).setCellValue("tch: ");
      s1.getRow(1).createCell(8).setCellValue(x2);
      if(x2>0)
        for(int i=0;i<x2;i++)
          for(int j=0;j<3;j++)
            s1.getRow(x1+1+i).createCell(10+j).setCellValue(conflictsDLP[i][j]);

      HSSFCellStyle cs=mf.fmt.formatStats();
      s1.getRow(0).getCell(7).setCellStyle(cs);
      s1.getRow(0).getCell(8).setCellStyle(cs);
      s1.getRow(1).getCell(7).setCellStyle(cs);
      s1.getRow(1).getCell(8).setCellStyle(cs);
    }
  }
}
