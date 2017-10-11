import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;

public class Window extends JFrame {
    public JTextField input;
    public JButton button;
    public JButton button2;
    public JButton button3;
    public JTextField output;
    public String tInput;
    public String tOutput;
    public String t;

    public JLabel Label;
    public JLabel Label1;
    public JLabel Label2;
    public JLabel Label3;
    public Window (){
        super("ExcelToExcel");
        setBounds(100, 100, 1200, 300);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        final Container con = getContentPane();
        con.setLayout(new FlowLayout());

        Label1 = new JLabel();
        button2 = new JButton("Обзор входного пути");
        button2.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int ret = fileChooser.showDialog(null, "Открыть файл");
                if (ret == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    input.setText(file.getAbsolutePath());
                }
            }
        });
        con.add(button2);
        Label1.setText("Выберите путь к папке входных файлов");//поле ввода
        con.add(Label1);
        input = new JTextField(100);//поле ввода
        con.add(input);
        Label2 = new JLabel();//поле ввода
        button3 = new JButton("Выходной файл");
        button3.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                JFileChooser fileChooser = new JFileChooser();
                int ret = fileChooser.showDialog(null, "Открыть файл");
                if (ret == JFileChooser.APPROVE_OPTION) {
                    File file = fileChooser.getSelectedFile();
                    output.setText(file.getAbsolutePath());
                }
            }
        });
        con.add(button3);
        Label2.setText("Введите путь и выходной файл в формате C:\\Users\\ApotinV\\123.xls");
        con.add(Label2);
        output = new JTextField(100);
        con.add(output, BorderLayout.WEST);
        button = new JButton("Поехали!");
        //button.setBounds(5, 5, 85, 30);
        con.add(button);
        Label3 = new JLabel();
        con.add(Label3);
        Label = new JLabel(t);
        con.add(Label, BorderLayout.EAST);

        button.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                tInput = input.getText();
                tOutput = output.getText();
                System.out.println(tInput);
                ArrayList<BkExcel> list1 = new ArrayList<BkExcel>();
                ArrayList<BkExcel> list2 = new ArrayList<BkExcel>();
                ArrayList<BkExcel> list = null;
                int countOfFiles = DirectoryFileNames.GetFileNames(tInput).size();
                for (int i = 0; i < countOfFiles; i++) {
                    System.out.println(DirectoryFileNames.GetFileNames(tInput).get(i));
                    try {
                        if (DirectoryFileNames.GetFileNames(tInput).get(i).
                                substring(DirectoryFileNames.GetFileNames(tInput).get(i).indexOf(".xl")).equals("xlsx"))
                        list = Parser.parseXlsx(tInput + "\\" + DirectoryFileNames.GetFileNames(tInput).get(i));
                        else list = Parser.parseXls(tInput + "\\" + DirectoryFileNames.GetFileNames(tInput).get(i));
                        if (i == 0) {
                            list2.add(list.get(0));
                            list2.add(list.get(1));
                        }
                        list1 = BkExcel.check(list/*, dateTake*/);
                        for (int j = 0; j < list1.size(); j++) {
                            list2.add(list1.get(j));
                        }
                    } catch (Exception e) {
                        Label3.setText(e.toString());
                        Label.setVisible(false);
                    }
                }
                BkExcel.writeIntoExcel(tOutput, list2);
                Label.setText("Excel файл успешно создан");
            }
        });
    }
}
