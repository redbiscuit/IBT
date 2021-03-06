package com.bj.sevenstar.basics;

import com.bj.requests.Requests;
import com.bj.requests.Session;
import com.bj.utils.RP;
import com.bj.utils.dataMgr.Data;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Hashtable;

/**
 * Created by JACKSON on 2017/8/30.
 */
public class InterfaceMock {

    public static Session getSession(){
        return Requests.session();
    }

    public static void processTask(ArrayList<Hashtable<String,String>> AL,String Env) throws IOException {
        Hashtable<String,String> HT = null;
        Session session = getSession();
        String result = null,URL = null;

        if(Env.equals("Backend")){
            URL = RP.getProperty("Backend_URL");
        }

        int i = 0;
        while(i<AL.size()){
            HT = AL.get(i);
            if(HT.get("Method").equals("Post")){
                result = session.post(URL+HT.get("URI")).forms(HT.get("Params")).send().readToText();
            }else if(HT.get("Method").equals("Get")){
                result = session.get(URL+HT.get("URI")).send().readToText();
            }
            System.out.println(result);
            i++;
        }
    }

    public static void main(String[] args) throws IOException {
        Data datapool = new Data();

        String URL = RP.getProperty("Backend_URL");

        datapool.setExeclInput("ST.xls");

        ArrayList<Hashtable<String,String>> AL = null;
        Hashtable<String,String> HT = null;

        AL = datapool.getData("设置","后台登录设置");
        processTask(AL,"Backend");
    }
}