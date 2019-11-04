import GC from "@grapecity/spread-sheets";
/**
 * sp常量
 * @create by Kellach 2019年10月22日
 */
export default class spjsconst{
  /**
   * 把JSON字符串反序列化成对象, 然后再用 fromJSON 来初始化 spread 对象
   */
  public static jsonOptions:any = {
    ignoreFormula: false, // indicate to ignore style when convert json to workbook, default value is false
    ignoreStyle: false, // indicate to ignore the formula when convert json to workbook, default value is false
    frozenColumnsAsRowHeaders: false, // indicate to treat the frozen columns as row headers when convert json to workbook, default value is false
    frozenRowsAsColumnHeaders: false, // indicate to treat the frozen rows as column headers when convert json to workbook, default value is false
    doNotRecalculateAfterLoad: false //  indicate to forbid recalculate after load the json, default value is false
 }
  /**
  * 把 spread toJSON 的返回的对象序列化成JSON字符串。
  */
  public static serializationOption:any = {
      ignoreStyle: false, // indicate to ignore the style when convert workbook to json, default value is false
      ignoreFormula: false, // indicate to ignore the formula when convert workbook to json, default value is false
      rowHeadersAsFrozenColumns: false, // indicate to treat the row headers as frozen columns when convert workbook to json, default value is false
      columnHeadersAsFrozenRows: false // indicate to treat the column headers as frozen rows when convert workbook to json, default value is false
  }

 /**
  * 绑定数据校验器
  * @create by Kellach 2019年10月22日
  * @param workbook
  */
 public static bindValitionAlert(workbook:GC.Spread.Sheets.Workbook):void{
    workbook.bind(GC.Spread.Sheets.Events.ValidationError, function(e:any, args:any) {
      var dv = args.validator;
      if (dv) {
          if (dv.showErrorMessage()) {
              var oldValue = args.sheet.getValue(args.row, args.col);
              var errorTitle = dv.errorTitle();
              var errorMessage = dv.errorMessage();
              var errorStyle = dv.errorStyle();
              if (errorStyle == GC.Spread.Sheets.DataValidation.ErrorStyle.stop) {
                  alert(errorMessage);
                  args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.retry;
              } else if (errorStyle == GC.Spread.Sheets.DataValidation.ErrorStyle.warning) {
                  var result = confirm(errorMessage);
                  if (result) {
                      args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.discard;
                  } else {
                      args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.retry;
                  }
              } else { //information
                  alert(errorMessage);
                  args.validationResult = GC.Spread.Sheets.DataValidation.DataValidationResult.forceApply;
              }
              console.log("title:"+errorTitle);
              console.log("message:"+errorMessage);
          }
      }
    });
 }

 /**
  * 绑定无限行列
  * @create by Kellach 2019年10月22日
  * @param workbook
  */
 public static bindMaxRowsAndCols(workbook:GC.Spread.Sheets.Workbook){
  workbook.bind(GC.Spread.Sheets.Events.ActiveSheetChanged, function (sender:any, args:any) {
    let sheet = args.newSheet;
    let rowCount = sheet.getRowCount();
    let bottomRow = sheet.getViewportBottomRow(1);
    let colCount = sheet.getColumnCount();
    let rightCol = sheet.getViewportRightColumn(1);
    if(rowCount<40){
      rowCount = 40;
      bottomRow = 41;
      sheet.setRowCount(rowCount);
    }
    if(colCount<40){
      colCount = 40;
      rightCol = 41;
      sheet.setColumnCount(colCount);
    }
    sheet.bind(GC.Spread.Sheets.Events.TopRowChanged, function (sender:any, args:any) {
      rowCount = sheet.getRowCount();
      bottomRow = sheet.getViewportBottomRow(1);
      if(bottomRow == rowCount-1){
        if(rowCount < 100000){
          sheet.setRowCount(rowCount+10);
        }
      }
    });
    sheet.bind(GC.Spread.Sheets.Events.LeftColumnChanged,function(sender:any, args:any){
      colCount = sheet.getColumnCount();
      rightCol = sheet.getViewportRightColumn(1);
      if(rightCol == colCount - 1){
        if(colCount < 1000){
          sheet.setColumnCount(colCount+10);
        }
      }
    });
  });
 }
}
