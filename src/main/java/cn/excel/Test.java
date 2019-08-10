package cn.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.sql.*;
import java.util.HashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Test {
    public static void main(String[] args) throws Exception {

        String regex="^(13[0-9]|14[579]|15[0-3,5-9]|16[6]|17[0135678]|18[0-9]|19[89])\\d{8}$";

        String phone="17812345678";
        if(phone.length()!=11){
            System.out.println("手机号应为十一位");
        }else{
            Pattern p = Pattern.compile(regex);
            Matcher m = p.matcher(phone);
            boolean isMatch = m.matches();
            if(isMatch){
               System.out.println("您的手机号" + phone + "是正确格式@——@");
            } else {
                System.out.println("您的手机号" + phone + "是错误格式！！！");
            }
        }



       /* Test tm = new Test();
        tm.jdbcex(true);*/
       /* f();*/

    }
    public static void f(){
        byte b[] = new byte[1024];
        try{

            FileInputStream fis = new FileInputStream("f:/d.xlsx");

            ProgressMonitorInputStream monitor = new ProgressMonitorInputStream(null,"读取文件",fis);
            int all = monitor.available();//整个文件的大小
            System.out.println(all);
            int in = monitor.read(b);//每次读取文件的大小
            System.out.println(in);
            int readed=0;//表示已经读取的文件
           /* readed+=in;//累加读取文件大小
            float process = (float)readed / all * 100;*/
            ProgressMonitor p=new ProgressMonitor(null,"读取文件",monitor.toString(),0,100);

            while(monitor.read(b)!=-1){
                readed+=in;//累加读取文件大小
                float process = (float)readed / all * 100;
                p.setNote("archived " + process + " %");// 显示在进度条上
                String s = new String(b);
               /* System.out.print(s);*/
                Thread.sleep(10);

            }
        }catch (Exception e) {
            e.printStackTrace();
        }

    }
    public void jdbcex(boolean isClose) throws InstantiationException, IllegalAccessException,
            ClassNotFoundException, SQLException, IOException, InterruptedException {
        String xlsFile = "f:/poiSXXFSBigData.xlsx";	//输出文件
//内存中只创建100个对象，写临时文件，当超过100条，就将内存中不用的对象释放。
        Workbook wb = new SXSSFWorkbook(100);	//关键语句
        Sheet sheet = null;	//工作表对象
        Row nRow = null;	//行对象
        Cell nCell = null;	//列对象
//使用jdbc链接数据库
        Class.forName("com.mysql.jdbc.Driver").newInstance();
        String url = "jdbc:mysql://192.168.0.222:3306/app_manage?characterEncoding=UTF-8";
        String user = "test123456";
        String password = "test123456";
//获取数据库连接
        Connection conn = DriverManager.getConnection(url, user,password);
        Statement stmt = conn.createStatement(ResultSet.TYPE_SCROLL_SENSITIVE,ResultSet.CONCUR_UPDATABLE);
        String sql = "select * from user limit 1000000"; //100万测试数据
        ResultSet rs = stmt.executeQuery(sql);
        ResultSetMetaData rsmd = rs.getMetaData();
        long startTime = System.currentTimeMillis();	//开始时间
        System.out.println("strat execute time: " + startTime);
        int rowNo = 0;	//总行号
        int pageRowNo = 0;	//页行号
        while(rs.next()) {
//打印300000条后切换到下个工作表，可根据需要自行拓展，2百万，3百万...数据一样操作，只要不超过1048576就可以
            if(rowNo%10000==0){
                System.out.println("Current Sheet:" + rowNo/10000);
                sheet = wb.createSheet("我的第"+(rowNo/10000)+"个工作簿");//建立新的sheet对象
                sheet = wb.getSheetAt(rowNo/10000);	//动态指定当前的工作表
                pageRowNo = 0;	//每当新建了工作表就将当前工作表的行号重置为0
            }
            rowNo++;
            nRow = sheet.createRow(pageRowNo++);	//新建行对象
// 打印每行，每行有6列数据 rsmd.getColumnCount()==6 --- 列属性的个数
            for(int j=0;j<rsmd.getColumnCount();j++){
                nCell = nRow.createCell(j);
                nCell.setCellValue(rs.getString(j+1));
            }
            if(rowNo%10000==0){
                System.out.println("row no: " + rowNo);
            }
//	Thread.sleep(1);	//休息一下，防止对CPU占用，其实影响不大
        }
        long finishedTime = System.currentTimeMillis();	//处理完成时间
        System.out.println("finished execute time: " + (finishedTime - startTime)/1000 + "m");
        FileOutputStream fOut = new FileOutputStream(xlsFile);
        wb.write(fOut);
        fOut.flush();	//刷新缓冲区
        fOut.close();
        long stopTime = System.currentTimeMillis();	//写文件时间
        System.out.println("write xlsx file time: " + (stopTime - startTime)/1000 + "m");
        if(isClose){
            this.close(rs, stmt, conn);
        }
    }
    //执行关闭流的操作
    private void close(ResultSet rs, Statement stmt, Connection conn ) throws SQLException{
        rs.close();
        stmt.close();
        conn.close();
    }
}
