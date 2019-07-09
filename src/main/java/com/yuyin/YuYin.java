package com.yuyin;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class YuYin {
    public static void main(String[] args) {
        ActiveXComponent activeXComponent=new ActiveXComponent("Sapi.SpVoice");
        try {
            //音量0-100
            activeXComponent.setProperty("Volume",new Variant(100));
            //语音朗读速度-10到+10
            activeXComponent.setProperty("Rate",new Variant(-2));
            //获取执行对象
            Dispatch dispatch=activeXComponent.getObject();
            //执行朗读
            Dispatch.call(dispatch,"Speak",new Variant("注意!来订单了，注意，来订单了。请及时查看"));
            //关闭执行对象
            dispatch.safeRelease();
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            //关闭应用程序链接
         activeXComponent.safeRelease();
        }
    }
}
