<template>
  <div class="firstQuickStart">
    <h1>This is an about quickStart</h1>
    <div class="btns_class_div">
      <button @click="addSheet">addSheet</button>
      <button @click="act">Test</button>
    </div>
    <div id = "workbookDiv" class="host_class"></div>
  </div>
</template>
<script lang="ts">
import { Component, Prop, Vue } from 'vue-property-decorator'
import element from 'element-ui'
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import '@grapecity/spread-sheets-vue'
import GC from "@grapecity/spread-sheets"

// GC.Spread.Common.CultureManager.culture("zh-cn");
/**
 * 模板前台
 * @create by Kellach 2019年9月26日
 */
@Component
export default class firstQuickStart extends Vue {

  private spread:any = null;

  /**
   * 周期函数
   */
  private mounted():void{
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById("workbookDiv"), { sheetCount: 1 });
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




}
</script>
<style lang="less">
  .host_class{
    width: 100%;
    height: 300px;
    border: 1px solid gray;
  }
  .btns_class_div{
    text-align: left;
    margin-bottom: 5px;
  }
</style>
