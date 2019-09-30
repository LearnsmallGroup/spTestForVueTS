import axios from '@/config/axios-config'
/**
 * Http 请求工具类
 * @update by Kellach 2019年7月17日 在请求中增加异常捕获
 */
class HttpUtils{

    /**
     * get请求 query
     * @param url
     * @param params
     */
    get(url:string,model = {}){
        return new Promise((resolve,reject)=>{
            axios.get(url,{
                params : model
            })
            .then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
            }).catch((err:any)=>{
                reject(err)
            })
        })
    }
    /**
     * post请求 insert
     * @param url
     * @param data
     */
    post(url:string,data={}){
        return new Promise((resolve,reject)=>{
            axios.post(url,JSON.stringify(data))
            .then((response:any) =>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
            })
            .catch((err:any)=>{
                reject(err);
            });
        })
    }
    /**
     * delete请求 删除
     * @param url
     * @param data
     */
    delete(url:string,data={}){
        return new Promise((resolve,reject)=>{
            axios.delete(url,data)
            .then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
            })
            .catch((err:any)=>{
                reject(err);
            });
        })
    }

    /**
     * put请求 修改
     * @param url
     * @param data
     */
    put(url:string,data={}){
        return new Promise((resolve,reject)=>{
            axios.put(url,JSON.stringify(data))
            .then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
            })
            .catch((err:any)=>{
                reject(err)
            })
        })
    }

    /**
     * 文件下载参数需要添加 config:responseType
     * @update 2019-03-07 by Kellach
     * @param url
     * @param model
     */
    downLoadGet(url:string,model = {}){
        return new Promise((resolve,reject)=>{
            axios.get(url,{
                params : model,
                responseType: 'blob'
            }).then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
                }).catch((err:any)=>{
                reject(err)
            })
        })
    }

    /**
     * 文件下载Post请求 参数需要添加 config:responseType
     * @update 2019-03-25 by Kellach
     * @param url
     * @param model
     */
    downLoadPost(url:string,data={}){
        return new Promise((resolve,reject)=>{
            axios.post(url,JSON.stringify(data),{
                responseType: 'blob'
            }).then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
                },(err:any)=>{
                    reject(err)
                });
        });
    }

    /**
     * 上传文件(为POST请求)
     * @create 2019年7月4日 by Han
     * @param url 路径
     * @param data 数据对象
     */
    uploadFiles(url:string,data:any){
        let param = new FormData();
        //通过append向form对象添加数据
        for(let key in data){
            param.append(key, data[key]);
        }

        let config = {
            //添加请求头
            headers: {
                "Content-Type": "multipart/form-data"
            }
        };
        return new Promise((resolve,reject)=>{
            axios.post(url,param,config)
            .then((response:any)=>{
                if(response == undefined){
                    let resp:any={
                        data:{msg:'网络请求失败！'}
                    }
                    reject(resp);
                }else{
                    resolve(response.data);
                }
            }).catch((err:any)=>{
                reject(err)
            })
        })
    }
    /**
     * 通用post方法，调用方自定义config
     * @param url
     * @param param
     * @param config
     * @Author Remon
     */
    postCommon(url:string,param:any,config:any){
        return new Promise((resolve,reject)=>{
            axios.post(url,param,config)
                .then((response:any)=>{
                    if(response == undefined){
                        let resp:any={
                            data:{msg:'网络请求失败！'}
                        }
                        reject(resp);
                    }else{
                        resolve(response.data);
                    }
                },(err:any)=>{
                    reject(err);
                });
        });
    }
}
export default new HttpUtils();
