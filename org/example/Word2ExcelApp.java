package org.example;

import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Word2ExcelApp {
    public static void main(String[] args) {
        Convert convert = new Convert();
        convert.initUI();
    }
}

class Convert extends JFrame {
    //word文件夹地址
    String wordUrl = "";
    //设置word文件夹路径文本框
    final JTextField textFieldWordFolder = new JTextField(50);
    //设置执行输出按钮
    final JButton buttonToExcel = new JButton("执行输出Excel");

    //创建界面
    public void initUI() {
        //设置标题
        this.setTitle("Word表格转Excel");
        //设置界面大小
        this.setSize(600, 140);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        //设置不可拉伸
        this.setResizable(false);
        //关闭流式布局
        this.setLayout(null);
        //开始界面元素设计
        //设置Word目录字样
        final JLabel labelWordDir = new JLabel("Word目录路径");
        //设置Word目录标签位置大小
        labelWordDir.setBounds(5, 5, 90, 25);
        //添加Word目录标签
        this.add(labelWordDir);

        //设置Word目录文本框位置大小
        textFieldWordFolder.setBounds(100, 5, 320, 25);
        //设置不可编辑
        textFieldWordFolder.setEditable(false);
        //设置Word目录文本框背景色为白色
        textFieldWordFolder.setBackground(Color.white);
        //设置Word目录文本框内容为空
        textFieldWordFolder.setText("");
        //添加Word目录文本框标签
        this.add(textFieldWordFolder);

        //设置获取Word文件夹目录按钮
        final JButton buttonWordDir = new JButton("获取Word文件夹位置");
        //设置Word文件夹位置大小
        buttonWordDir.setBounds(430, 5, 150, 25);
        //添加Word文件夹按钮
        this.add(buttonWordDir);
        //设置输出Excel按钮位置
        buttonToExcel.setBounds(430, 40, 130, 50);
        //添加输出Excel按钮
        this.add(buttonToExcel);
        //结束界面元素设计

        //设置按钮监听器
        ButtonListener buttonListener = new ButtonListener(this);
        //将获取Word文件夹位置按钮添加监听
        buttonWordDir.addActionListener(buttonListener);
        //将转换Excel按钮添加监听
        buttonToExcel.addActionListener(buttonListener);
        //设置窗体可见
        this.setVisible(true);
    }

    //打开Word文件夹对话框
    public void showFolderOpenDialog() {
        //System.out.println("按键点击测试 -> 打开文件夹");
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File(""));
        //只允许打开文件夹类型
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        //不允许多选
        chooser.setMultiSelectionEnabled(false);
        int result = chooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            wordUrl = chooser.getSelectedFile().getAbsolutePath();
            System.out.println(wordUrl.replaceAll("\\\\", "\\\\\\\\") + " -> showFileOpenDialog()");
            textFieldWordFolder.setText(wordUrl);
        }
    }

    //执行转换
    public void ToExcel() {
        //System.out.println("按键点击测试 -> 执行转换");
        if (wordUrl.equals("")) {
            JOptionPane.showMessageDialog(null, "请选择文件夹", null, JOptionPane.WARNING_MESSAGE);
        } else {
            //将地址转义
            String wordUrl1 = wordUrl.replaceAll("\\\\", "\\\\\\\\") + "\\\\\\\\";

            //根据SimpleDateFormat类生成日期，并以日期为名字生成excel文件
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd-HH-mm-ss");
            String date = simpleDateFormat.format(new Date());
            //String excelPath = "src/main/java/org/example/" + date + ".xlsx";
            String excelPath = wordUrl1 + date + ".xlsx";
            System.out.println(excelPath);
            //定义Excel表格
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet("sheet1");
            //如果要实现下方的每隔N次换行的话，这里要先创建第一行
            XSSFRow row = sheet.createRow(0);
            try {
                FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
                workbook.write(fileOutputStream);
                fileOutputStream.flush();
                fileOutputStream.close();
                System.out.println("新建Excel文件成功");
            } catch (Exception e) {
                System.out.println("新建Excel文件失败");
                e.printStackTrace();
            }

            //实例化Document
            Document spireDocument = new Document();
            //声明ArrayList存放文件
            ArrayList<String> arrayList = new ArrayList<String>();
            //实例化File类并传入文档所在文件夹
            File resourcesFolder = new File(wordUrl1);
            //扫描指定文件夹里的文件
            File[] resourcesFile = resourcesFolder.listFiles();
            //存放结果集
            ArrayList<String> resultList = new ArrayList<String>();
            for (File file : resourcesFile) {
                if (file.getName().equals("无效文件.docx") || file.getName().equals("无效文件.doc")) {
                    continue;
                }
                if (file.getName().endsWith(".doc") || file.getName().endsWith(".docx")) {
                    arrayList.add(file.getName());
                }
            }

            for (int getList = 0; getList < arrayList.size(); getList++) {
                String filePath = wordUrl1 + arrayList.get(getList);
                String filePathDocx = "";
                System.out.println("******************** START " + arrayList.get(getList) + " ********************");
                try {
                    System.out.println("******************** 开始转换格式 " + arrayList.get(getList) + " ********************");
                    if (filePath.endsWith(".doc")) {
                        spireDocument.loadFromFile(filePath);
                        filePathDocx = filePath.replace(".doc", ".docx");
                        spireDocument.saveToFile(filePathDocx, FileFormat.Docx);
                    }
                    if (filePath.endsWith(".docx")){
                        filePathDocx=filePath;
                    }
                    System.out.println("******************** 结束转换格式 " + arrayList.get(getList) + " ********************");
                    FileInputStream fileInputStream = new FileInputStream(new File(filePathDocx));
                    XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
                    //获取文档中的全部表格
                    List<XWPFTable> tables = xwpfDocument.getTables();
                    List<XWPFTableRow> rows;
                    List<XWPFTableCell> cells;
                    for (XWPFTable table : tables) {
                        //获取表格对应的行
                        rows = table.getRows();
                        for (XWPFTableRow xwpfTableRow : rows) {
                            //获取行对应的单元格
                            cells = xwpfTableRow.getTableCells();
                            for (XWPFTableCell xwpfTableCell : cells) {
                                //System.out.println(xwpfTableCell.getText());
                                resultList.add(xwpfTableCell.getText().trim());
                            }
                        }
                    }
                    fileInputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                System.out.println("******************** END " + arrayList.get(getList) + " ********************");
            }

            try {
                int countRow = 0;
                for (int i = 0; i < resultList.size(); i++) {
                    if ((i != 0) && (i % 12 == 0)) {
                        countRow++;
                        row = sheet.createRow(countRow);
                    }
                    String value = resultList.get(i);
                    row.createCell(i % 12).setCellValue(value);
                }

                FileOutputStream fileOutputStream = new FileOutputStream(excelPath);
                fileOutputStream.flush();
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                System.out.println("更新Excel文件成功");
            } catch (Exception e) {
                System.out.println("更新Excel文件失败");
                e.printStackTrace();
            }
            JOptionPane.showMessageDialog(null, "成功输出!", null, JOptionPane.INFORMATION_MESSAGE);
            System.exit(0);
        }
    }
}

class ButtonListener implements ActionListener {
    private Convert convert;

    public ButtonListener(Convert convert) {
        this.convert = convert;
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        if (e.getActionCommand().equals("获取Word文件夹位置")) {
            convert.showFolderOpenDialog();
        }
        if (e.getActionCommand().equals("执行输出Excel")) {
            convert.ToExcel();
        }
    }
}