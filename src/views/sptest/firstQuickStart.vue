<template>
  <div class="firstQuickStart">
    <h1>This is an about quickStart</h1>
    <div class="btns_class_div">
      <button @click="addSheet">addSheet</button>
      <button @click="upfile">uploadFile</button>
      <button @click="getFileData">获取数据</button>
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
            action="http://127.0.0.1:11221/dealExcel/uploadFile"
            :limit="1"
            accept=".xlsx"
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
   * 周期函数
   */
  private mounted():void{
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById("workbookDiv"), { sheetCount: 1 });
    var fbx = new GC.Spread.Sheets.FormulaTextBox.FormulaTextBox(document.querySelector('#formulaBar') as Object,{});
    fbx.workbook(this.spread);
    var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(document.querySelector('#statusBar') as HTMLElement,{});
    statusBar.bind(this.spread);
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
    this.coverObj = this.$loading(this.coverOptions);
    let jsonOptions:any = {
       ignoreFormula: false, // indicate to ignore style when convert json to workbook, default value is false
       ignoreStyle: false, // indicate to ignore the formula when convert json to workbook, default value is false
       frozenColumnsAsRowHeaders: false, // indicate to treat the frozen columns as row headers when convert json to workbook, default value is false
       frozenRowsAsColumnHeaders: false, // indicate to treat the frozen rows as column headers when convert json to workbook, default value is false
       doNotRecalculateAfterLoad: false //  indicate to forbid recalculate after load the json, default value is false
    }
    ajax.get("/dealExcel/getModelFile/Test01")
      .then((response:any)=>{
        let div:any = document.getElementById('workbookDiv');
        let spread:any = GC.Spread.Sheets.findControl(div);
        spread.suspendPaint();
        // let fileJson:string = JSON.stringify(response.model);
        let sjon:string = JSON.parse(response.model);
        spread.fromJSON(sjon, jsonOptions);
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

      }).finally(()=>{
        this.coverObj.close();
      });
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
