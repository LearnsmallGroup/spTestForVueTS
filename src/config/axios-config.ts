import vue from "vue";
import axios from "axios";
import router from '@/router';
import elm from 'element-ui';
import Qs from 'qs';


//配置超时时间
axios.defaults.timeout = 1000*60*10;

//设置基础访问路径
axios.defaults.baseURL="http://127.0.0.1:11221";
axios.defaults.withCredentials=true;
axios.interceptors.request.use(
    (config:any) => {
        let token: string = "";
        let flag =true;
        for(var a in config.headers) {
          if(a == 'Content-Type'){
            flag=false;
            break;
          }
        }
        flag && (config.headers = (() => {
            let baseHead: any = {
                'Content-Type': 'application/json'
            }
            return baseHead;
        })());
        //重写参数序列化函数
        config.paramsSerializer = (param:any):string => {
             return Qs.stringify(param, {arrayFormat: 'brackets'})
        }
        //调用新的序列化函数
        config.paramsSerializer(config.params);
        return config;
    },
    (error:any) => {
        return Promise.reject(error);
    }
)

axios.interceptors.response.use(
    (response:any)=>{
        return response;
    },(err:any) =>{
        if(!err.response){
            elm.Message({
                message: err.message,
                type: 'error'
            });
        }else{
            return Promise.reject(err.response)
        }
    }
)
export default axios;
