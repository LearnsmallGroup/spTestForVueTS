<template>
  <div class="testModel">
    <h1>测试模板</h1>
    <div class="btns_class_div">
      <button @click="getFileData">获取数据</button>
      <button @click="getExcelFileData">获取后台Excel数据</button>
      <button @click="getFileEmptyData">获取空模板数据</button>
      <button @click="reCal">恢复计算</button>
      <button @click="getCurrentSheet">从服务器获取当前sheet内容</button>
      <button @click="getTagData">获取Tag数据渲染</button>
      <button @click="testCycleTime">测试遍历所有cell用时</button>
      <button @click="getMessageCount">查询当前校验错误的数量</button>
      <button @click="lockSheetName">锁定解锁sheetName</button>
      <button @click="getGCwb">获取后台拼接的workBook</button>
      <button @click="getDataFormular">获取公式依赖</button>
    </div>
    <div id="formulaBar"  contenteditable="true" spellcheck="false" style="font-family: Calibri;border: 1px solid #808080;width:100%;height:35px;background:white;font-size: x-large ;"></div>
    <div id = "workbookDiv" class="host_class"></div>
    <div id="statusBar"></div>
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
import spconst from '@/common/spjsconst';


var ExcelIO = require("@grapecity/spread-excelio");
GC.Spread.Common.CultureManager.culture("zh-cn");

/**
 * 测试模板
 * @create by Kellach 2019年10月22日
 */
@Component
export default class testModel extends Vue {
  //遮罩层对象 --- start ----
  private coverOptions:any=COVER.COVER_OPTIONS;
  private coverObj:any = null;
  //遮罩层对象 --- end ----
  private workbook: any = null;
   /**
   * 周期函数
   */
  private mounted():void{
    this.workbook = new GC.Spread.Sheets.Workbook(document.getElementById("workbookDiv"), { sheetCount: 1 });
    var fbx = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(document.querySelector('#formulaBar') as Object,{});
    fbx.workbook(this.workbook);
    var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(document.querySelector('#statusBar') as HTMLElement,{});
    statusBar.bind(this.workbook);
    spconst.bindMaxRowsAndCols(this.workbook);
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
          ajax.get("/dealExcel/getModelFile/"+ms.value)
          .then((response:any)=>{
            let div:any = document.getElementById('workbookDiv');
            let workbook:any = GC.Spread.Sheets.findControl(div);
            let start = (new Date()).getTime();
            workbook.suspendPaint();
            workbook.suspendCalcService();
            let str:string = response.model;
            workbook.fromJSON(str,spconst.jsonOptions);
            let csource:any = response.source;
            spconst.bindValitionAlert(workbook);
            workbook.resumePaint();
            let end = (new Date()).getTime();
            console.log('用时：'+(end-start) +" 毫秒");
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
   * 获取后台GC渲染的Excel数据
   * @create by Kellach 2019年10月30日
   */
  public getExcelFileData():void{
    this.$prompt('请输入模板CODE', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消'
        }).then((ms:any) => {
          this.coverObj = this.$loading(this.coverOptions);
          ajax.get("/dealExcel/getExcelFile/"+ms.value)
          .then((response:any)=>{
            let div:any = document.getElementById('workbookDiv');
            let workbook:any = GC.Spread.Sheets.findControl(div);
            let start = (new Date()).getTime();
            workbook.suspendPaint();
            workbook.clearSheets();
            workbook.suspendCalcService();
            let str:string = response.model;
            workbook.fromJSON(str,spconst.jsonOptions);
            let csource:any = response.source;
            spconst.bindValitionAlert(workbook);
            workbook.resumePaint();
            let end = (new Date()).getTime();
            console.log('用时：'+(end-start) +" 毫秒");
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
   * 获取空模板数据
   * @create by Kellach 2019年10月22日
   */
  public getFileEmptyData():void{
    let cworkbook = this.workbook;
    this.$prompt('请输入模板CODE', '提示', {
      confirmButtonText: '确定',
      cancelButtonText: '取消'
    }).then((ms:any) => {
        this.coverObj = this.$loading(this.coverOptions);
          ajax.get("/dealExcel/getEmptyFile/"+ms.value)
          .then((response:any)=>{
            cworkbook.suspendPaint();
            let str:string = response.model;
            cworkbook.fromJSON(str,spconst.jsonOptions);
            let csource:any = response.source;
            cworkbook.resumePaint();
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
   * 重新恢复sheet间的计算功能
   * @create by Kellach 2019年10月22日
   */
  private reCal():void{
    let div:any = document.getElementById('workbookDiv');
    let workbook:any = GC.Spread.Sheets.findControl(div);
    workbook.resumeCalcService();
    workbook.refresh();
  }
  /**
   * 获取当前sheet页的Json
   * @create by Kellach 2019年10月22日
   */
  private getCurrentSheet():void{
    let cworkbook = this.workbook;
    let sheet:any = cworkbook.getActiveSheet();
    let sheetName:string = sheet.name();
    let modelCode:string = "sheetTest";
    this.coverObj = this.$loading(this.coverOptions);
    ajax.get("/dealExcel/getModelSheetFile/"+modelCode+"/"+sheetName)
      .then((response:any)=>{
        // cworkbook.suspendPaint();
        let str:string = response.model;
        sheet.fromJSON(str);
        // cworkbook.resumePaint();
      })
      .catch((error:any)=>{
        console.log(error);
      })
      .finally(()=>{
        this.coverObj.close();
      });
  }
  /**
   * 获取tag数据渲染
   * @create by Kellach 2019年10月22日
   */
  private getTagData():void{
    this.$prompt('请输入模板CODE', '提示', {
          confirmButtonText: '确定',
          cancelButtonText: '取消'
        }).then((ms:any) => {
          this.coverObj = this.$loading(this.coverOptions);
          ajax.get("/dealExcel/getModelFile/"+ms.value)
          .then((response:any)=>{
            //1.渲染模板
            let div:any = document.getElementById('workbookDiv');
            let workbook:any = GC.Spread.Sheets.findControl(div);
            workbook.suspendPaint();
            let str:string = response.model;
            workbook.fromJSON(str,spconst.jsonOptions);
            let csource:any = response.source;
            workbook.resumePaint();
            debugger;
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
    测试遍历所有cell用时
    @create by Kellach 2019年10月23日
   */
  private testCycleTime():void{
    let start = (new Date()).getTime();
    let sheetcount:number = this.workbook.getSheetCount();
    let i = 0;
    for(i;i<sheetcount;i++){
      let sheet:any = this.workbook.getSheet(i);
      let colcount:number = sheet.getColumnCount();
      let rowcount:number = sheet.getRowCount();
      for(let j=0;j<rowcount;j++){
        for(let k = 0;k<colcount;k++){
          sheet.getCell(j,k).tag();
         // console.log(sheet.name()+"第 "+j+" 行"+"第"+k+"列");
        }
      }
    }
    let end = (new Date()).getTime();
    alert('用时：'+(end-start) +" 毫秒");
  }
  /**
   * 数据校验数量
   * @create by Kellach 2019年10月23日
   */
  private getMessageCount():void{
    let count = 0;
    debugger;
    let sheet:any = this.workbook.getActiveSheet();
    let colcount:number = sheet.getColumnCount();
      let rowcount:number = sheet.getRowCount();
      for(let j=0;j<rowcount;j++){
        for(let k = 0;k<colcount;k++){
          let cell = sheet.getCell(j,k);
          if(cell != undefined && cell.validator() != undefined){
            let dv = cell.validator();
            let flag:boolean = dv.isValid(dv,j,k,cell.value());
            if(!flag){
              console.log(cell.value()+":"+ flag);
              count++;
            }
          }
        }
      }
      alert(count);
  }
  private islock:boolean = false;
  /**
   * 锁定解锁sheetName 修改
   * @create by Kellach 2019年10月23日
   */
  private lockSheetName():void{
    if(!this.islock){
      alert("锁定");
    }else{
      alert("解锁");
    }
    this.workbook.options.tabEditable = this.islock;
    this.islock = !this.islock;
  }
  /**
   * 从后台获取后台拼接sheet的workbook
   * @create by Kellach 2019年10月23日
   */
  private getGCwb():void{
    this.coverObj = this.$loading(this.coverOptions);
    ajax.get("/dealExcel/getGCWb/sheetTest")
    .then((response:any)=>{
      //1.渲染模板
      let div:any = document.getElementById('workbookDiv');
      let workbook:any = GC.Spread.Sheets.findControl(div);
      workbook.suspendPaint();
      let str:string = response.model;
      workbook.fromJSON(str,spconst.jsonOptions);
      let csource:any = response.source;
      workbook.resumePaint();
      debugger;
    }).catch((error:any)=>{
        console.log(error);
    }).finally(()=>{
      this.coverObj.close();
    });
  }
  /**
   * 获取公式依赖
   * @create by Kellach 2019年10月31日
   */
  private getDataFormular():void{
    let sheet:any = this.workbook.getActiveSheet();
    debugger;
  }
}
</script>
<style lang="less">
.host_class{
    width: 100%;
    height: 650px;
    border: 1px solid gray;
  }
  .btns_class_div{
    text-align: left;
    margin-bottom: 5px;
  }
</style>
