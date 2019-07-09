package com.yuyin;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.*;
/**
 * Txt class
 *
 * @author yuxiang
 * @since 2019-5-31
 * */
public class Txt {
    public static void main(String[] args) throws IOException {
        ActiveXComponent sap = new ActiveXComponent("Sapi.SpVoice");
        //输入文件
        FileInputStream srcFile = new FileInputStream("C:/Users/Administrator/Desktop/123.txt");
        InputStreamReader inputStreamReader=new InputStreamReader(srcFile,"UTF-8");
        //使用包装字符流读取文件
        BufferedReader br = new BufferedReader(inputStreamReader);
        String content = br.readLine();
        try {
        // 音量 0-100
            sap.setProperty("Volume", new Variant(100));
        // 语音朗读速度 -10 到 +10
            sap.setProperty("Rate", new Variant(-1));
        // 获取执行对象
            Dispatch sapo = sap.getObject();
        // 执行朗读
            while (content != null) {
                Dispatch.call(sapo, "Speak", new Variant(content));
                content = br.readLine();
            }
        // 关闭执行对象
            sapo.safeRelease();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            srcFile.close();
            inputStreamReader.close();
            br.close();
        // 关闭应用程序连接
            sap.safeRelease();
        }
    }
}
