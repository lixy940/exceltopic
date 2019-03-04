package main;

import utils.ExcelSAXParserUtil;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author LIS
 * @date 2019/3/4 15:20
 */
public class ExcelToPicUtils
{

    private JFrame frm;
    private JButton open;
    private JButton export;
    private JPanel p;
    private File source;
    private File target;
    private JFileChooser fc;
    private JFileChooser fcT;
    private int flag;
    private JTextArea t;
    private JTextArea h;
    public ExcelToPicUtils()
    {
        frm=new JFrame("Excel成绩单图片导出--陈小凤专属");
        open=new JButton("打开");
        export =new JButton("导出");


        fc=new JFileChooser();
        fc.setAcceptAllFileFilterUsed(false);// 取消所有文件过滤项
        fc.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fc.setFileFilter(new FileNameExtensionFilter("Excel文件", "xls","xlsx"));// 设置只过滤扩展名为.xls的Excel文件

        fcT=new JFileChooser();
        fcT.setAcceptAllFileFilterUsed(false);// 取消所有文件过滤项
        fcT.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        Container c=frm.getContentPane();
        FlowLayout flowLayout = new FlowLayout();
        flowLayout.setAlignment(FlowLayout.CENTER);
        c.setLayout(flowLayout);
        JLabel label1 = new JLabel("                                                                                                                                                                                                                                                                    ");
        label1.setOpaque(true);
        c.add(label1);
        p=new JPanel();
        p.add(open);
        p.add(export);
        c.add(p);
        label1 = new JLabel("                                                                                                                                                                                                                                                                    ");
        label1.setOpaque(true);
        c.add(label1);
        JLabel label = new JLabel("输入含有名字的列序号(从1开始为第一列):");
        label.setOpaque(true);
        c.add(label);
        t=new JTextArea(1,6);
        t.setToolTipText("输入数字，如0，1，2,...");
        t.setText(null);
        c.add(t);
        label1 = new JLabel("                                                                                                                                                                                                                                                                    ");
        label1.setOpaque(true);
        c.add(label1);
        label = new JLabel("   设置输出的列的宽度，如150 200 250等:");
        label.setOpaque(true);
        c.add(label);
        h=new JTextArea(1,8);
        h.setText("200");
        c.add(h);
        label1 = new JLabel("                                                                                                                                                                                                                                                                    ");
        label1.setOpaque(true);
        c.add(label1);
        label = new JLabel("步骤：1.首先点击‘打开’按钮，选择excel文件；2.输入含有名字的列序号;3.设置输出每列的宽度;4.点击‘导出’按钮选择图片输出目录，并点击保存按钮即可");
        label.setOpaque(true);
        c.add(label);

        //注册按钮事件
        open.addActionListener(new Action());
        export.addActionListener(new Action());
        Toolkit tk= Toolkit.getDefaultToolkit();
        Dimension d=tk.getScreenSize();
        frm.setSize((int)d.getWidth(),(int)d.getHeight());
        frm.setSize(900,300);
        frm.setVisible(true);
        //设置默认的关闭操作
        frm.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
    }


    public void openFile()  //打开文件
    {
        //设置打开文件对话框的标题
        fc.setDialogTitle("打开文件");

        //这里显示打开文件的对话框
        try{
            flag=fc.showOpenDialog(frm);
        } catch(HeadlessException head){
            System.out.println("Open File Dialog ERROR!");
        }

        //如果按下确定按钮，则获得该文件。
        if(flag== JFileChooser.APPROVE_OPTION)
        {
            //获得该文件
            source=fc.getSelectedFile();
            System.out.println("open file----"+source.getName());
        }
    }


    private void exportFile()//保存文件
    {
        String fileName;
        //设置保存文件对话框的标题
        fcT.setDialogTitle("设置保存目录");
        //这里将显示保存文件的对话框
        try{
            flag=fcT.showSaveDialog(frm);
        }
        catch(HeadlessException he){
            System.out.println("Save Directory Dialog ERROR!");
        }

        //如果按下确定按钮，则获得该文件。
        if(flag== JFileChooser.APPROVE_OPTION)
        {
            //获得你输入要保存的文件
            target=fcT.getSelectedFile();
            //获得文件名
            fileName = target.getName();
            String indexStr = t.getText();
            String widthStr = h.getText();
            try {
                Integer.parseInt(indexStr);
            } catch (Exception e1) {
                JOptionPane.showMessageDialog(null, "序号输入不对，只能为数字","提示消息", JOptionPane.WARNING_MESSAGE);
                return;
            }
            try {
                Integer.parseInt(widthStr);
            } catch (Exception e1) {
                JOptionPane.showMessageDialog(null, "宽度设置不正确，只能为正整数","提示消息", JOptionPane.WARNING_MESSAGE);
                return;
            }
            try {
                handlerExcel(source.getAbsolutePath(), Integer.valueOf(t.getText()), target.getAbsolutePath(), Integer.parseInt(widthStr));
                JOptionPane.showMessageDialog(null, "图片导出成功，图片存储在目录【" + target.getAbsolutePath() + "】下","提示消息", JOptionPane.INFORMATION_MESSAGE);
                System.exit(0);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(null, "图片导出失败："+e.getMessage(),"提示消息", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    public static void handlerExcel(String path, Integer column, String outPutPath, Integer singleWidth)throws Exception {
        ExcelSAXParserUtil saxParserUtil = new ExcelSAXParserUtil();
        saxParserUtil.processOneSheet(path);

        List<String> dataList = saxParserUtil.getDataList();

        int width = dataList.get(0).split(" ").length*singleWidth;
        int height = 64*2;
        Font font = new Font("微软雅黑", Font.PLAIN, 25);

        int len = dataList.size();

        Map<String, Integer> nameMap = new HashMap<String, Integer>();

        System.out.println("正在生成，请稍等......");
        System.out.println();
        for (int i = 1; i < len; i++) {
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

            String[] headArr = dataList.get(0).split(" ");

            for (int j = 0; j < headArr.length; j++) {
                g.drawString("————————————",  j * singleWidth, y-16);
                g.drawString(headArr[j],40+j*singleWidth,y+5);
                if (j > 0 && j < headArr.length) {
                    g.drawString("|",j*singleWidth,y);
                }
            }

            y = clip.height - descent -44;
            for (int j = 0; j <headArr.length ; j++) {
                g.drawString("————————————",  j * singleWidth, y);
                if (j > 0 && j < headArr.length) {
                    g.drawString("|",j*singleWidth,y);
                }

            }

            y = clip.height - descent -18;
            String[] dataArr = dataList.get(i).split(" ");

            for (int j = 0; j < dataArr.length; j++) {
                g.drawString(dataArr[j],40+j*singleWidth,y-5);
                if (j > 0 && j < dataArr.length) {
                    g.drawString("|",j*singleWidth,y);
                }
                g.drawString("————————————",  j * singleWidth, y+13);
            }

            //如果有重复名字
            String name = dataArr[column - 1];
            if (!nameMap.containsKey(name)) {
                nameMap.put(name, 0);
            }else{
                nameMap.put(name, nameMap.get(name) + 1);
                name = name+nameMap.get(name);
            }
            File outFile = new File( outPutPath+"\\" + name + ".png");

            g.dispose();
            ImageIO.write(image, "png", outFile);// 输出png图片

        }
        System.out.println("导出数据成功，数据存储在目录：" + outPutPath + "下");
    }
    //按钮监听器类内部类
    class Action implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {

            //判断是哪个按纽被点击了
            if(e.getSource()==open) {
                openFile();
            } else if(e.getSource()== export) {
                exportFile();
            }
        }
    }

    public static void main(String[] args)throws Exception
    {
        UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        new ExcelToPicUtils();

    }
}
