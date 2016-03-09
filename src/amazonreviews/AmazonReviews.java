/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package amazonreviews;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author shubham_15294
 */
public class AmazonReviews {

    /**
     * @param args
     */
    static ArrayList<String> reviews;
    static ArrayList<String> reviewsDate;
    static ArrayList<String> reviewsStar;
    static String name;

    static String getUrlSource(String url) throws IOException {
        //System.setProperty("http.proxyHost", "172.31.1.4");
        //System.setProperty("http.proxyPort", "8080");
        URL yahoo = new URL(url);
        URLConnection yc = yahoo.openConnection();
        BufferedReader in = new BufferedReader(new InputStreamReader(yc.getInputStream(), "UTF-8"));
        String inputLine;
        StringBuilder a = new StringBuilder();
        while ((inputLine = in.readLine()) != null) {
            a.append(inputLine);
            a.append("\n");
        }
        in.close();

        return a.toString();
    }

    static void GetReviews(String PID, String pageNo) throws IOException {
        String url = "http://www.amazon.com/product-reviews/" + PID
                + "/?ie=UTF8&showViewpoints=0&pageNumber=" + pageNo
                + "&sortBy=bySubmissionDateDescending";
        //String url = "http://www.amazon.in/product-reviews/" + PID + "/ref=cm_cr_dp_see_all_summary?ie=UTF8&showViewpoints=" + pageNo + "&sortBy=byRankDescending";
        System.out.println(url);
        String source = "";
        try {
            source = getUrlSource(url);
        } catch (Exception e) {

        }
        String source1;
        Pattern pattern = Pattern.compile("<span class=\"a-size-base review-text\">(.*?)</span>");
        Matcher matcher = pattern.matcher(source);
        Pattern pattern1 = Pattern.compile("<title>Amazon.com: Customer Reviews: (.*?)</title>");
        Matcher matcher1 = pattern1.matcher(source);
        Pattern pattern2 = Pattern.compile("<span class=\"a-size-base a-color-secondary review-date\">on(.*?)</span>");
        Matcher matcher2 = pattern2.matcher(source);
        Pattern pattern3 = Pattern.compile("<span class=\"a-size-base review-title a-text-bold\">(.*?)</span>");
        Matcher matcher3 = pattern3.matcher(source);
        Pattern pattern4 = Pattern.compile("<span class=\"a-icon-alt\">(.*?)out of 5 stars</span>");
        Matcher matcher4 = pattern4.matcher(source);
        while (matcher.find()) {
            String r = matcher.group(1);
            r = r.replace(",", "");
            reviews.add(r);
        }
        while (matcher1.find()) {
            name = matcher1.group(1);
        } while (matcher2.find()) {
            String reviewDate = matcher2.group(1);
            reviewDate = reviewDate.replace(",", "");
            reviewsDate.add(reviewDate);
        } while (matcher4.find()) {
            String reviewStars = matcher4.group(1);
            reviewStars = reviewStars.replace(",", "");
            //System.out.println(reviewTitle);
            reviewsStar.add(reviewStars);
            //System.out.println("entering"); 
        }

    } 
    
    public static void main(String[] args) throws IOException {
        // TODO Auto-generated method stub
        //new AmazonReviews.filea("B00I8BIBCW");
        String s1 = "B002RL9CYK";
        reviews = new ArrayList<String>();
        reviewsDate = new ArrayList<String>();
        reviewsStar = new ArrayList<String>();
        XSSFWorkbook workbook = new XSSFWorkbook();
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");
          
        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("0", new Object[] {"Review Text", "Review Date", "Review Stars"});
        for (int i = 1; i <= 100; i++) {
            GetReviews(s1, Integer.toString(i));
        }
        for (int i = 0; i < reviews.size(); i++) {
            data.put(Integer.toString(i + 1), new Object[] {reviews.get(i), reviewsDate.get(i), reviewsStar.get(i)});  
        }
        
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            XSSFRow row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File(name + ".xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }
}
