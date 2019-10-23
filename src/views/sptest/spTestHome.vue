<template>
  <div class="spTestHome">
    <h1>This is an about SpreadJS demo for practice</h1>
    <div id="nav">
      <router-link to="/spTest">spTest</router-link> |
      <router-link to="/quickStart">firstQuickStart</router-link> |
      <router-link to="/testModel">模板测试</router-link>
    </div>
  </div>
</template>
<script lang="ts">
import { Component, Prop, Vue } from 'vue-property-decorator'
import element from 'element-ui'

/**
 * 练习入口文件
 * @create by Kellach 2019年9月26日
 */
@Component
export default class spTestHome extends Vue {

}
</script>
<style lang="less">
#nav {
  padding: 30px;
  a {
    font-weight: bold;
    color: #2c3e50;
    &.router-link-exact-active {
      color: #42b983;
    }
  }
}
</style>
