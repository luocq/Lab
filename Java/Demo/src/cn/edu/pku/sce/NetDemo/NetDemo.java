package cn.edu.pku.sce.NetDemo;

import java.net.InetAddress;
import java.net.UnknownHostException;

public class NetDemo {
	public static void main(String args[]){
        try{ //以下代码通过域名建立InetAddress对象：
            InetAddress addr = InetAddress.getByName("www.sce.pku.edu.cn");
            String domainName = addr.getHostName();//获得主机名
            String IPName = addr.getHostAddress();//获得IP地址
            System.out.println(domainName);
            System.out.println(IPName);
        }catch(UnknownHostException e){
            e.printStackTrace();
        }
    }
}
