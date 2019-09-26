<template>
  <div class="spTest">
    <h1>This is an about spTest</h1>
    <div>
      <gc-spread-sheets :hostClass='hostClass' >
        <gc-worksheet
          :dataSource="dataTable"
          :autoGenerateColumns = 'autoGenerateColumns'
        >
          <gc-column
            :width="width"
            :dataField="'price'"
            :visible = 'visible'
            :formatter = 'formatter'
            :resizable = 'resizable'
          ></gc-column>
        </gc-worksheet>
      </gc-spread-sheets>
    </div>
  </div>
</template>
<script lang="ts">

import { Component, Prop, Vue } from 'vue-property-decorator'
import element from 'element-ui'
import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
import '@grapecity/spread-sheets-vue'
import GC from "@grapecity/spread-sheets"

GC.Spread.Common.CultureManager.culture("zh-cn");
/**
 * 做测试使用
 * @create by Kellach 2019年9月24日
 */
@Component
export default class spTest extends Vue {

     private hostClass:any = 'spread-host';
     private autoGenerateColumns:boolean = true;
     private width:number =300;
     private visible:boolean = true;
     private resizable:boolean = true;
     private formatter:any = "$ #.00";
     private spread:any = null;

  /**
   * 声明周期钩子函数
   */
  private mounted():void{
    let spdiv:any = document.getElementById("spTest");
  }
  /**
   * 初始化数据
   */
  get dataTable():Array<any>{
    let dataTable:Array<any> = [];
    for (let i:number = 0; i < 42; i++) {
      let pis:number = i + 0.56;
      let pram:any = {price: pis};
      dataTable.push(pram)
    }
    return dataTable;
  }
}
</script>

<style lang="less">
.spread-host {
    height: 600px;
}
</style>
