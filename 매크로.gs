/** @OnlyCurrentDoc */
function myFunction() {
  var ss = SpreadsheetApp.getActive();
  var mySheet1 = ss.getSheetByName('주보고');
  var mySheet2 = ss.getSheetByName('월보고');
  var mySheet3 = ss.getSheetByName('사업보고');
  
  var lastRow1 = mySheet1.getLastRow();
  var lastRow2 = mySheet2.getLastRow();
  var lastRow3 = mySheet3.getLastRow();
  var lastCol = mySheet1.getLastColumn();
 
  var temp = [];
   
 
 
 //1. 한달이 총 몇주(week) 인지, 몇월(month)인지  확인하기,  
  var month_last_row = lastRow1 ;                            // 예) 4월이 끝나자마자 월 합산을 할 때를 대비해,  월 마지막주 로우를 엑셀 마지막 로우로 초기화 
  for (var i=0; i<=8; i++) {
    temp.push(mySheet1.getRange(lastRow1-i,3).getValue());   // 월 행 거슬러서 올라가며 데이터 넣기 temp[0]=마지막데이터  , temp[1]= 마지막-1 ROW의 데이터   
    if(i>=1) {                                               // 월 데이터가 두개이상 쌓이면 확인    
      if  (temp[i-1]!=temp[i]) {                             // 아래행과 현재 행이 월이 다르면 확인 
        if (i<=3) {month_last_row = lastRow1-i } ;           // 예) 4월데이터까지 쌓인 다음 3월 보고서 작성할 때를 대비하여, 3주를 지나지 않고 월이바뀌면 3월의 마지막 주 ROW정보만 저장함 
        if (i >3) {                                  
          var month_start_row = lastRow1-i+1;                // 4주 이상으로 구성되었으면 해당월의 시작주가 있는 ROW정보 저장 
          var week = month_last_row - month_start_row +1;    // WEEK수 계산    
          var month = temp[i-1];                             // 몇월인지  
          var year = mySheet1.getRange(month_start_row,2).getValue();  //년도 저장 
          console.log('month: ' +month+ ' start : '+month_start_row+ ' last : '+ month_last_row+' week : '+week);   //테스트 
          
          break;
        };  
      };
    };
  };
  
  
  
  // 2.  출석간부수~ 기타전 까지 월 합계 넣기 
  value=[]; 
  var col_num1 = getByName(mySheet1,'출석 [간부수]');   //출석간부수의 열 이름 가진 열의 열번호 찾기 
  var col_num2 = getByName(mySheet1,'기타');

  for (var j=col_num1 ; j< col_num2; j++) {            //  해당월의 출석 간부수 ~ 기타전 열까지 다음 반복하기(기타는 따로 계산)   
    var myRange1 = mySheet1.getRange(month_start_row,j,week,1);  // (월 시작 주 행, 출석간부수 열, 주수행, 열은 확장안함) 각 열마다 주수만큼 범위설정  
    var value = myRange1.getValues(); // 각 열의 월 처음 주 부터 마지막 주 까지의 값 저장  
    var sum1=0;

    for (var i=0 ; i<=week-1;i++ ) {        
       sum1 = sum1 + +value[i];                    // string을 number로 바꾸어 더하기 
    };
    
    //console.log('sum1: ' + sum1);
    mySheet2.getRange(2,j).setValue(sum1);             //월보고시트에 저장  (월보고는 2번째 열에 덮어쓰기 ) 
    mySheet3.getRange(lastRow3+1,j).setValue(sum1);    //사업보고시트에 저장  ( 사업보고는 마지막 행 다음행에 새로 쓰기 )
  };


  //3. '기타'열은 문자로 더하기   
  value=[]; 
  
  var myRange2 = mySheet1.getRange(month_start_row,col_num2,week,1);  //기타열의 월 처음 주부터 마지막 주 까지 범위 선택 
  var sum2='';
  var value = myRange2.getValues();
  
  for (var i=0 ; i<=week-1;i++ ) {
     sum2=sum2+value[i];
  };
  mySheet3.getRange(lastRow3+1,col_num2).setValue(sum2);  //사업보고서 기타값에만 저장 
  
  
//4. 회계보고 (전 주 내용을 가져옴 ) 
 var col_num4 = getByName(mySheet1,'잔액');            // 주보고 잔액열의 열번호 찾기 
 
 for (var j=col_num2+1 ; j<col_num4; j++) {            // 해당월의 기타열 다음(이번달 비밀헌금) ~ 잔액 전 열까지 다음 반복하기(잔액은 따로 계산)   
    var myRange1 = mySheet1.getRange(month_start_row+1,j,week,1);  // 각 열의 월 2번째 주행 부터 다음 달 첫번째주행 까지 범위 선택 (회계보고는 한주씩 늦음) 
    var sum1=0;
    var value = myRange1.getValues(); // 회계보고 열의 월 두번째 주부터 다음달 첫번째 주 까지 범위 선택
    
    for (var i=0 ; i<=week-1;i++ ) {
       sum1 = sum1 + +value[i];                    // string을 number로 바꾸어 더하기 
    };
    mySheet2.getRange(2,j).setValue(sum1);             //월보고시트 회계열에 저장  
    
    mySheet3.getRange(lastRow3+1,j).setValue(sum1);    //사업보고시트에 저장 
    
  };
 


 //5. 지난달 잔액 
 var fillDownRange = mySheet3.getRange(lastRow3+1,col_num4);       
     mySheet3.getRange(lastRow3,col_num4).copyTo(fillDownRange);    //사업보고 잔액:  지난 달 잔액수식을 복사해옴  
 
  mySheet2.getRange(2,col_num4).setValue(mySheet3.getRange(lastRow3+1,col_num4).getValue()); //월보고서 잔액 : 사업보고서에서 가져옴 
  mySheet2.getRange(2,col_num4+1).setValue(mySheet3.getRange(lastRow3,col_num4).getValue()); // 월례보고 지난달이월금 : 사업보고서 지난달 잔액을 가져옴  
  
 mySheet2.getRange(2,col_num2+1,1,6).setNumberFormat('#,##0');                               // 셀서식 : 콤마  
 mySheet3.getRange(lastRow3+1,col_num2+1,1,5).setNumberFormat('#,##0');

 //6. 그 밖의 항목 (년, 월, 마지막, 시작회차  )
  var col_num3 = getByName(mySheet1,'회차');
  var last_hoicha = mySheet1.getRange(month_last_row,col_num3).getValue();
  
  mySheet2.getRange(2,2).setValue(year);
  mySheet2.getRange(2,3).setValue(month);
  mySheet2.getRange(2,4).setValue(last_hoicha-week+1);   //1열에 시작회차넣기 
  mySheet2.getRange(2,5).setValue(last_hoicha);          //회차 열에 마지막회차 넣기 
 
  mySheet3.getRange(lastRow3+1,2).setValue(year);
  mySheet3.getRange(lastRow3+1,3).setValue(month+'월');
  mySheet3.getRange(lastRow3+1,4).setValue(last_hoicha-week+1);   //1열에 시작회차넣기 
  mySheet3.getRange(lastRow3+1,5).setValue(last_hoicha);          //회차 열에 마지막회차 넣기 
 
 
};


//특정 시트, 특정열의 이름으로 열번호 찾는 function 
  function getByName(sheet, colName) {
    var data = sheet.getRange("A1:1").getValues();  //1행의 모든 value를 저장=> data[0]에 저장됨  
    var col = data[0].indexOf(colName)+1;            //index는 0부터 시작하므로, +1 
    return col;
  };

//행 삽입하는 function 
function insert_row(month_last_row) {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange((month_last_row+1) + ':'+ (month_last_row+1)).activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  
};

//셀서식
function myFunction1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('CR2:CW2').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('#,##0');
};