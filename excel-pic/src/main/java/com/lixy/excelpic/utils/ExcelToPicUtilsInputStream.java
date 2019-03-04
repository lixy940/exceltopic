package com.lixy.excelpic.utils;

/**
 * @Author: MR LIS
 * @Description:
 * @Date: Create in 11:15 2018/7/6
 * @Modified By:
 */

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.util.List;

public class ExcelToPicUtilsInputStream {
    private static String path = "D:\\考试成绩.xlsx";

    public static void main(String[] args) throws Exception{

        while (true) {
            System.out.println("请将excel文件放在--> D盘 根目录下，并取名为 ‘考试成绩.xlsx’");
            System.out.println("==========================================================================");
            System.out.println("指明excel中含有名字的列是第几列，输入数字即可，如第一列输入数字1，然后回车");
            System.out.println();
            System.out.print("请在此输入，并回车：");
            BufferedReader br = new BufferedReader(new InputStreamReader(System.in));
            try {
                String s = br.readLine();
                int column = Integer.parseInt(s);
                getExcelData(path, column);
                break;
            } catch (Exception e) {
                System.out.println();
                System.out.println("输入不合法，请输入数字");
                System.out.println();
                System.out.println("==========================================================================");
                System.out.println();
                System.out.println();
            }
        }
    }

    public static void getExcelData(String path, Integer column) throws Exception {
        ExcelSAXParserUtil saxParserUtil = new ExcelSAXParserUtil();

        saxParserUtil.processOneSheet(path);
        List<String> spitDataList = saxParserUtil.getDataList();

        String outPutDir = path.substring(0, path.lastIndexOf(File.separator) + 1) + "pic";
        File dir = new File(outPutDir);
        if (!dir.exists()) {
            dir.mkdirs();
        }

        //数据列数量
        int len = spitDataList.get(0).split(" ").length;

        int width = len * 150;
        int height = 2 * 64;

        Font font = new Font("黑体", Font.BOLD, 25);

        //从列头后的一行开始
        for (int i = 1; i < spitDataList.size(); i++) {
            // 创建图片
            BufferedImage image = new BufferedImage(width, height,
                    BufferedImage.TYPE_INT_BGR);
            Graphics g = image.getGraphics();

            g.setClip(0, 0, width, height);
            g.setColor(Color.white);
            g.fillRect(0, 0, width, height);// 先用黑色填充整张图片,也就是背景
            g.setColor(Color.black);// 在换成黑色
            g.setFont(font);// 设置画笔字体
            /** 用于获得垂直居中y */
            Rectangle clip = g.getClipBounds();
            FontMetrics fm = g.getFontMetrics(font);
            int ascent = fm.getAscent();
            int descent = fm.getDescent();


            int y = (clip.height - (ascent + descent)) / 4 + ascent;

            String[] headArr = spitDataList.get(0).split(" ");
            for (int j = 0; j < headArr.length; j++) {
                // 画出字符串
                g.drawString(headArr[j], 40 + j * 150, y);
            }

            for (int j = 0; j < len; j++) {
                g.drawString("--------------------------------------------", j * 150, y + 25);// 画出字符串
            }

            y = clip.height - descent - 22;
            String[] dataArr = spitDataList.get(i).split(" ");
            for (int j = 0; j < dataArr.length; j++) {
                // 画出字符串
                g.drawString(dataArr[j], 40 + j * 150, y);
            }

            File outFile = new File(outPutDir + "\\" + spitDataList.get(i).split(" ")[column] + ".png");
            g.dispose();
            // 输出png图片
            ImageIO.write(image, "png", outFile);
        }
        System.out.println();
        System.out.println("=====================================================================");
        System.out.println("导出数据成功，数据存储在目录：" + outPutDir + "下");


    }

}

