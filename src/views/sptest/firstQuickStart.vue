<template>
  <div class="firstQuickStart">
    <h1>This is an about quickStart</h1>
    <div class="btns_class_div">
      <button @click="addSheet">addSheet</button>
      <button @click="upfile">uploadSSjson</button>
      <button @click="getFileData">获取数据</button>
      <button @click="submitData">提交数据</button>
      <button @click="changeLanguage">切换中英文</button>
      <input type='file' @change="processFile($event)"/>
      <button @click="upExcel">前端导入</button>
      <button @click="getLocaltionIndex">获取坐标</button>
    </div>
    <div id="formulaBar"  contenteditable="true" spellcheck="false" style="font-family: Calibri;border: 1px solid #808080;width:100%;height:35px;background:white;font-size: x-large ;"></div>
    <div id = "workbookDiv" class="host_class"></div>
    <div id="statusBar"></div>
    <div >
      <el-dialog
          title="上传文件"
          :visible.sync="upVisiable"
          width="30%"
          append-to-body
          :close-on-click-modal = false
          :before-close="upfileBeforeClose"
          >
          <el-upload
            class="upload-demo"
            ref="upload"
            action="http://127.0.0.1:11221/dealExcel/uploadFileBySheet"
            :limit="1"
            accept=".ssjson"
            :auto-upload="false"
            :file-list="fileList">
            <el-button slot="trigger" size="small" type="primary">点击上传</el-button>
            <el-button style="margin-left: 10px;" size="small" type="success" @click="submitUpload">上传到服务器</el-button>
          </el-upload>
      </el-dialog>
    </div>
  </div>
</template>
<script lang="ts">
import { Component, Prop, Vue } from 'vue-property-decorator'
import ElementUI from 'element-ui';
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import '@grapecity/spread-sheets-vue'
import "@grapecity/spread-sheets-resources-zh"
import GC from "@grapecity/spread-sheets"
import ajax from '@/config/HttpUtils';
import COVER from '@/config/LoadingCover';
import pako from 'pako';

var ExcelIO = require("@grapecity/spread-excelio");
GC.Spread.Common.CultureManager.culture("zh-cn");
/**
 * 模板前台
 * @create by Kellach 2019年9月26日
 */
@Component
export default class firstQuickStart extends Vue {

  //遮罩层对象 --- start ----
    private coverOptions:any=COVER.COVER_OPTIONS;
    private coverObj:any = null;
    //遮罩层对象 --- end ----
  private spread:any = null;
  /**
   * 把JSON字符串反序列化成对象, 然后再用 fromJSON 来初始化 spread 对象
   */
  private jsonOptions:any = {
       ignoreFormula: false, // indicate to ignore style when convert json to workbook, default value is false
       ignoreStyle: false, // indicate to ignore the formula when convert json to workbook, default value is false
       frozenColumnsAsRowHeaders: false, // indicate to treat the frozen columns as row headers when convert json to workbook, default value is false
       frozenRowsAsColumnHeaders: false, // indicate to treat the frozen rows as column headers when convert json to workbook, default value is false
       doNotRecalculateAfterLoad: false //  indicate to forbid recalculate after load the json, default value is false
    }
  /**
   * 把 spread toJSON 的返回的对象序列化成JSON字符串。
   */
  private serializationOption:any = {
       ignoreStyle: false, // indicate to ignore the style when convert workbook to json, default value is false
       ignoreFormula: false, // indicate to ignore the formula when convert workbook to json, default value is false
       rowHeadersAsFrozenColumns: false, // indicate to treat the row headers as frozen columns when convert workbook to json, default value is false
       columnHeadersAsFrozenRows: false // indicate to treat the column headers as frozen rows when convert workbook to json, default value is false
    }

  /**
   * 周期函数
   */
  private mounted():void{
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById("workbookDiv"), { sheetCount: 1 });
    var fbx = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(document.querySelector('#formulaBar') as Object,{});
    fbx.workbook(this.spread);
    var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(document.querySelector('#statusBar') as HTMLElement,{});
    statusBar.bind(this.spread);
    this.setMenuDatas();
  }
  /**
   * 添加Sheet
   * @create by Kellach 2019年9月26日
   */
  private addSheet():void{
    //直接添加新Sheet
    //  this.spread.addSheet(this.spread.getSheetCount());
    //添加带名字的sheet
    /**
     * 这里的sheetName 是不能重复的 否则会报错
     */
    let count:number = this.spread.getSheetCount();
    let sheet:any =new GC.Spread.Sheets.Worksheet("new Sheet"+count);
    this.spread.addSheet(count,sheet);

  }
  /**
   * 激活某个sheet
   */
  private setSheetActive(idx:number):void{
    this.spread.setActiveSheetIndex(idx);
  }
  /**
   * 获取表格中被激活的Sheet
   */
  private getActiveSheet():any{
    return this.spread.getActiveSheet();
  }
  private upVisiable:boolean = false;
  private fileList:Array<any>=[];
  /**
   * 上传文件
   * @create by Kellach 2019年9月29日
   */
  private upfile():void{
    this.upVisiable = true;
  }
  /**
   * 文件上传关闭时
   */
  private upfileBeforeClose():void{
    this.upVisiable = false;
  }
  private submitUpload():void{
    this.upVisiable = false;
    (this.$refs.upload as HTMLFormElement).submit();
  }

  /**
   * 获取数据文件
   */
  private getFileData():void{
    this.$prompt('请输入模板CODE', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消'
        }).then((ms:any) => {
          this.coverObj = this.$loading(this.coverOptions);
          // ajax.get("/dealExcel/getModelFile/"+ms.value)
          ajax.get("/dealExcel/getModelFile/"+ms.value)
          .then((response:any)=>{
            let div:any = document.getElementById('workbookDiv');
            let spread:any = GC.Spread.Sheets.findControl(div);
            spread.suspendPaint();
            // let fileJson:string = JSON.stringify(response.model);
            let str:string = response.model;
            // let sjon:any = JSON.parse(JSON.stringify(str));
            // spread.fromJSON(sjon, this.jsonOptions);
            spread.fromJSON(str,this.jsonOptions);
            // let source:string = JSON.stringify(response.source);
            // console.log(source);
            //前端渲染
            spread.resumePaint();
            // let csource:any = JSON.parse(response.source);
            let csource:any = response.source;
            //获取table
            let sheet:any = spread.getActiveSheet();
            let table:any = sheet.tables.findByName("gcTable0");
            table.bindingPath('zysh');
            let source = new GC.Spread.Sheets.Bindings.CellBindingSource(csource);
            sheet.setDataSource(source);
            table.showHeader(false);
            spread.resumePaint();
          }).catch((error:any)=>{
              console.log(error);
          }).finally(()=>{
            this.coverObj.close();
          });
        }).catch(() => {
          this.$message({
            type: 'info',
            message: '取消输入'
          });
        });

  }
  /**
   * 提交数据
   * @create by Kellach 2019年10月9日
   */
  private submitData():void{
      let value :any = this.findCellByTagName("data").value();
      alert(value);

      // let jsons:string = JSON.stringify(spread.toJSON(this.serializationOption));
      // this.coverObj = this.$loading(this.coverOptions);
      // ajax.post("/dealExcel/getJsonData",{json:jsons})
      // .then((response:any)=>{

      // })
      // .catch((error:any)=>{
      //   console.log(error);
      // })
      // .finally(()=>{
      //   this.coverObj.close();
      // });

  }
  /**
   * 通过tagName获取cell
   */
  public findCellByTagName(tagName:any):any{
      let sheet:any = this.getActiveSheet();
      let spreadNS:any = GC.Spread.Sheets;
      let condition:any = new spreadNS.Search.SearchCondition();
      condition.searchTarget = spreadNS.Search.SearchFoundFlags.cellTag;
      condition.searchString = tagName;
      condition.searchOrder = spreadNS.Search.SearchOrder.zOrder;
      let result:any = sheet.search(condition);
      let col:number = result.foundColumnIndex;
      let row:number = result.foundRowIndex;
      if(col==-1||row==-1){
        return null;
      }else{
        return sheet.getCell(row,col);
      }
  }
  private cstyle:any = undefined;
  /**
   * 添加右键菜单
   * @create by Kellach 2019年10月9日
   */
  public setMenuDatas():void{
    let unLockCell:any = {
            text: "打标",
            name: "tagFlag",
            command: "tagFlagFunction",
            workArea: "viewport",
            // subMenu: [
            //     {
            //         name: "selectColorPicker",
            //         command: "selectWithBg"
            //     }
            // ]
        };
    this.spread.contextMenu.menuData.push(unLockCell);

    let commandManager:any = this.spread.commandManager();
    let tagFlagFunction:any={
        canUndo: true,
        execute: function (spread:any, options:any) {


          let sheet:any = spread.getActiveSheet();
          let style:any = new GC.Spread.Sheets.Style();

          sheet.suspendPaint();

          //删除上一次打标记录
          let spreadNS:any = GC.Spread.Sheets;
          let condition:any = new spreadNS.Search.SearchCondition();
          condition.searchTarget = spreadNS.Search.SearchFoundFlags.cellTag;
          condition.searchString = "dataYYY";
          condition.searchOrder = spreadNS.Search.SearchOrder.zOrder;
          let result:any = sheet.search(condition);
          let col:number = result.foundColumnIndex;
          let row:number = result.foundRowIndex;
          let cell:any = sheet.getCell(row,col);
          if(col!=-1 && row!=-1){
            cell.tag(null);
            //还原背景色，添加表格线
            var lineStyle = GC.Spread.Sheets.LineStyle.thin;
            var lineBorder = new GC.Spread.Sheets.LineBorder( '#D4D4D4',lineStyle);
            var sheetArea = GC.Spread.Sheets.SheetArea.viewport;
            sheet.getRange(row,col,1, 1).setBorder(lineBorder, {outline:true}, sheetArea);
            cell.backColor('rgb(255,255,255)');
          }
          //从新打标
          style.backColor = 'rgb(255,0,0)';
          let selections:any = sheet.getSelections();
          let selectionIndex:number = 0, selectionCount:number = selections.length;
          for (; selectionIndex < selectionCount; selectionIndex++) {
              let selection:any = selections[selectionIndex];
              for (let i:number = selection.row; i < (selection.row + selection.rowCount); i++) {
                  for (let j:number = selection.col; j < (selection.col + selection.colCount); j++) {
                      if(selection.rowCount>1||selection.colCount>1){
                        alert("只能选中一个单元格！");
                        sheet.resumePaint();
                        return;
                      }
                      sheet.setStyle(i, j, style, GC.Spread.Sheets.SheetArea.viewport);
                      sheet.getCell(i, j).tag('dataYYY');
                  }
              }
          }
          sheet.resumePaint();
        }
    };
    commandManager.register("tagFlagFunction", tagFlagFunction, null, false, false, false, false);
  }
  private lg:boolean = true;
  /**
   * 切换中英文
   * @create by Kellach 2019年10月10日
   */
  public changeLanguage():void{
    let sheet:any = this.getActiveSheet();
    this.spread.suspendPaint();
    let colNum:number = sheet.getColumnCount();
    let rowNum:number = sheet.getRowCount();
    for(let i = 0;i<rowNum;i++){
      for(let j =0;j<colNum;j++){
        let tag = sheet.getCell(i,j).tag();
        if(tag!=undefined){
          let tmp:string = JSON.stringify(tag);
          let obj:any = eval ("(" + JSON.parse(tmp) + ")");
          if(this.lg){
            sheet.getCell(i,j).value(obj.en);
          }else{
            sheet.getCell(i,j).value(obj.zh);
          }
        }
      }
    }
    this.lg = !this.lg;
    this.spread.resumePaint();
  }
  //前台导入相关
  private excelFile:any = {};
  processFile (event:any) {
    this.excelFile = event.target.files[0];
  }
  public upExcel():void{
    var self = this;
    let sheet:any = this.getActiveSheet();
    this.spread.suspendPaint();
    // if(this.excelFile.name.substring(this.excelFile.name.lastIndexOf(".")+1) == "ssjson"){
      if(true){
      var reader = new FileReader();
      reader.readAsText(this.excelFile);
      reader.onload = function () {

        var obj:any = this.result;
        console.log("文件内容：");
        console.log(escape(obj));
        console.log("压缩前");
        console.log(obj.length);
        var binaryString = pako.gzip(obj,{ to: 'string' });
        console.log("压缩内容：");
        console.log(binaryString);

        // let code = encodeURI(binaryString);
        // console.log("编码大小："+code.length);
        let config = {
            //添加请求头
            // headers: {
            //     "Content-Type": "multipart/form-data"
            // }
        };
        // ajax.postCommon("/test/upfile",binaryString,config)
         let data:any = new FormData();
         data.append('file',binaryString);
        ajax.postCommon("/test/upfile",data,config)
          .then((response:any)=>{

          }).catch((error:any)=>{
            console.log(error);
          }).finally(()=>{

          });

        // console.log("压缩后");
        // console.log(binaryString.length);
        // console.log("还原后")
        // var charData = binaryString.split('').map(function (x:any) { return x.charCodeAt(0); });
        // var data:any = pako.inflate(new Uint8Array(charData))

        // var res = '';
        // var chunk = 8 * 1024;
        // var i;
        // for (i = 0; i < data.length / chunk; i++) {
        //   res += String.fromCharCode.apply(null, data.slice(i * chunk, (i + 1) * chunk));
        // }
        // res += String.fromCharCode.apply(null, data.slice(i * chunk));
        // let restr = unescape(res);
        // console.log(restr.length);
        // console.log("还原内容：\n"+restr);





        // var obj:any = this.result;
        self.spread.fromJSON(JSON.parse(obj));
        self.spread.resumePaint();
      }
    }else{
      var excelIO = new ExcelIO.IO();
      excelIO.open(this.excelFile, function(json:string) {
        self.spread.fromJSON(json,self.jsonOptions);
        self.spread.resumePaint();
      });

    }
    /**
     * 导入数据的时候，需要绑定校验器
     */
    self.spread.bind(GC.Spread.Sheets.Events.ValidationError, function(e:any, args:any) {
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
          }
      }
    });
  }
  /**
   * 初始化坐标函数
   * @create by Kellach 2019年10月18日
   */
  private getLocaltionIndex():void{
    let sheet:any = this.getActiveSheet();
    this.spread.suspendPaint();
    let that:any = this;
    sheet.setText(0, 0, "Click anywhere inside of this Spread instance.");
    let top = this.spread.getHost().offsetTop;
    let left = this.spread.getHost().offsetLeft;
    this.spread.getHost().addEventListener('click', function(e:any){
        var y = e.pageY - top;
        var x = e.pageX - left;
        var result = that.spread.hitTest(x, y);
        var str = that.getHitAreaName(result);
        if (result) {
            let res = "(x: " + result.x.toString() + "," + " y:" + result.y.toString() + ")" + " ; " + str;
            alert(res);
        }
    });
    this.spread.resumePaint();
  }
  /**
   * 获取区域名称
   */
  private getHitAreaName(result:any):any{
    let str:any = "";
    if(result) {
        var worksheetHitInfo = result.worksheetHitInfo;
        var tabStripHitInfo = result.tabStripHitInfo;
        if (worksheetHitInfo) {
            var hitTestType = worksheetHitInfo.hitTestType;
            if (hitTestType === 0) {
                str = 'corner';
            } else if (hitTestType === 1) {
                str = 'colHeader';
            } else if (hitTestType === 2) {
                str = 'rowHeader';
            } else {
                var row = worksheetHitInfo.row;
                var col = worksheetHitInfo.col;
                str = 'viewport; ' + '(row: ' + row + ', col: ' + col + ')';
            }
        } else if (tabStripHitInfo) {
            if (tabStripHitInfo.navButton) {
                str = tabStripHitInfo.navButton;
            } else if (tabStripHitInfo.sheetTab) {
                str = tabStripHitInfo.sheetTab.sheetName;
            } else if (tabStripHitInfo.resize === true) {
                str = "resize";
            } else {
                str = "blank";
            }
        } else if (result.horizontalScrollBarHitInfo) {
            str = result.horizontalScrollBarHitInfo.element;
        } else if (result.verticalScrollBarHitInfo) {
            str = result.verticalScrollBarHitInfo.element;
        } else if (result.footerCornerHitInfo) {
            str = result.footerCornerHitInfo.element;
        }
    }
    return str;
  }

}
</script>
<style lang="less">
  .host_class{
    width: 100%;
    height: 800px;
    border: 1px solid gray;
  }
  .btns_class_div{
    text-align: left;
    margin-bottom: 5px;
  }
</style>
