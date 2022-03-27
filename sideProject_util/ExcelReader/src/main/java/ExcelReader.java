import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import org.junit.Test;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.*;

public class ExcelReader {

    // Excel絕對路徑 (支援xls)
    private final static String EXCELPATH = "C:\\Users\\Tibame_T14\\Desktop\\project\\test.xls";
    private final static String tableName = "test1"; // table名字
    // jdbc 相關資訊
    private final static String URL = "jdbc:mysql://localhost:3306/EXCEL_TEST?useUnicode=yes&characterEncoding=utf8&useSSL=true&serverTimezone=Asia/Taipei";
    ;
    public final static String USER = "root";
    public final static String PASSWORD = "password";

    @Test
    public void insertIntoDB() {
        //取得Map 存放 每個row的data
        Map<Integer, List<String>> data = readData(EXCELPATH);
        // 建立list 存放jdbc setMethod
        List<String> method = new ArrayList<>();


        // 取得屬性
        StringBuffer str = splitString(data, 0);
        // 有多少屬性
        int numOfAttr = getNumOfAttr(data);
        // 多少屬性 就有多少問號
        StringBuffer qusMark = new StringBuffer("( ");
        for (int i = 0; i < numOfAttr; i++) {
            qusMark.append(" ? , ");
        }
        qusMark.replace(qusMark.lastIndexOf(","), qusMark.length(), " ) ");

        // 準備jdbc setMethod();
        for (String methodName : data.get(1)) {
            method.add("set" + methodName);
        }


        // 做字串切割
        String insertSql = "insert into " + tableName + " (" + str
                + " ) values " + qusMark + ";";
        System.out.println(insertSql);

        // jdbc
        Connection connection = null;
        PreparedStatement ps = null;
        try {
            connection = DriverManager.getConnection(URL, USER, PASSWORD);
            ps = connection.prepareStatement(insertSql);
            connection.setAutoCommit(false);
            // 從data.get(2) 開始才裝資料  外層跑row
            for (int i = 2; i < data.size(); i++) {
                // 內層跑col
                for (int j = 0; j < numOfAttr; j++) {
                    switch (method.get(j)) {
                        case "setString":
                            ps.setString(j + 1, data.get(i).get(j));
                            break;
                        case "setBoolean":
                            ps.setBoolean(j + 1, Boolean.parseBoolean(data.get(i).get(j)));
                            break;
                        case "setInteger":
                            ps.setInt(j + 1, Integer.parseInt(data.get(i).get(j)));
                            break;
                        case "setDouble":
                            ps.setDouble(j + 1, Double.parseDouble(data.get(i).get(j)));
                            break;
                        case "setBLOB":
                            try{
                                File file = new File(data.get(i).get(j));
                               FileInputStream fileInputStream = new FileInputStream(file);
                                BufferedInputStream bis = new BufferedInputStream(fileInputStream);
                                byte[] bytes = new byte[bis.available()];
                                bis.read(bytes);
                                ps.setBytes(j+1,bytes);
                            }catch (Exception e){
                                System.out.println("找不到檔案");
                            }
                            break;
                    }
                }
                int res = ps.executeUpdate();
                System.out.println(res);
            }
            connection.commit();
        } catch (Exception e) {
            e.printStackTrace();
            try {
                connection.rollback();
            } catch (SQLException throwables) {
                throwables.printStackTrace();
            }
        } finally {
            try {
                ps.close();
                connection.close();
            } catch (SQLException throwables) {
                throwables.printStackTrace();
            }

        }
    }

    public static StringBuffer splitString(Map<Integer, List<String>> data, int rowNum) {
        // 取得attr 字串切割取得表格屬性
        List<String> attr = data.get(rowNum);
        // 裝入表格屬性
        StringBuffer str = new StringBuffer(" ");
        // 存入的data
        StringBuffer instValue = new StringBuffer();
        for (String s : attr) {
            str.append(s + " ,");
        }
        // 去除最後兩個,
        str.delete(str.length() - 2, str.length());
//        System.out.println(str);
        return str;
    }

    public static int getNumOfAttr(Map<Integer, List<String>> data) {
        List<String> l = data.get(0);
        return l.size();
    }

    /**
     * 工具 : 用來讀取excel
     * 取得一個map
     * index 0 存table的屬性
     * index 1 存放schema
     * index 2 之後存資料
     */
    public Map<Integer, List<String>> readData(String excelPath) {
        // 用來裝attr
        List<String> attr = new ArrayList<>();
        Map<Integer, List<String>> data = new HashMap<>();


        try {
            //取得目標
            File file = new File(excelPath);
            //讀取excel檔
            // 取得 workSheet
            Workbook workbook = Workbook.getWorkbook(file);
            // 得到第一個sheet
            Sheet sheet = workbook.getSheet(0);

            // 取得第一行 (也就是表格的屬性)
            Cell[] row = sheet.getRow(0);
            for (Cell c : row) {
                String contents = c.getContents();
                attr.add(contents);
            }

            /**
             *  每筆資料取得
             */
            // 先獲得筆數
            int nums_data = sheet.getRows();

            for (int i = 0; i < nums_data - 1; i++) {
                Cell[] eachData = sheet.getRow(i);
                // 建立一個arr 接收一個row的資料
                // row.length 是屬性的數量
                List<String> collectData = new ArrayList<>();
                for (Cell c : eachData) {
                    // 取得每隔data
                    String contents = c.getContents().trim();
                    if (contents != null && contents != "") {
                        collectData.add(contents);
                    }
                }
                data.put(i, collectData);
//                System.out.println();
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }

        // 檢查map有無取得資料
//        Set<Integer> keySet = data.keySet();
//        for (int i :keySet) {
//            List<String> d = data.get(i);
//            System.out.println(d.toString());
//        }

        return data;
    }

    @Test
    public void test() {
        Map<Integer, List<String>> integerListMap = readData(EXCELPATH);
    }
}
