package com.owoez;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by owoez on 7/22/2020
 *
 * @author : owoez
 * @date : 7/22/2020
 * @project : Excel Converter
 */
public class Filer extends JFrame{
  private JPanel filePanel;
  private JTextField filename;
  private JButton selectFile;
  private JButton compute;
  private JTextArea processInformationArea;
  private String absolutePath;

  public Filer(String title){
    super();
    this.setDefaultCloseOperation(EXIT_ON_CLOSE);
    this.setContentPane(filePanel);
    this.pack();
    processInformationArea.setText("No file selected yet");
    selectFile.addActionListener(new ActionListener() {
      @Override
      public void actionPerformed(ActionEvent e) {
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        int returnValue = jfc.showOpenDialog(null);
        // int returnValue = jfc.showSaveDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
          File selectedFile = jfc.getSelectedFile();
          String nameOfFile = selectedFile.getName();
          absolutePath = selectedFile.getAbsolutePath();
          filename.setText(nameOfFile);
          processInformationArea.setText(nameOfFile + " selected....");
          System.out.println(selectedFile.getAbsolutePath());
        }
      }
    });


    compute.addActionListener(new ActionListener() {
      @Override
      public void actionPerformed(ActionEvent e) {
        try {
          processFile(absolutePath);
        } catch (IOException ioException) {
          ioException.printStackTrace();
        } catch (InvalidFormatException invalidFormatException) {
          invalidFormatException.printStackTrace();
        }
      }
    });
  }

  public void processFile(String absolutePath) throws IOException, InvalidFormatException {
    processInformationArea.setText(processInformationArea.getText() + "\nReading file.........");
    ArrayList allUsers = new ArrayList<String>();
    ArrayList usersWithCourse = new ArrayList<String>();
    try {
      FileInputStream file = new FileInputStream(absolutePath);
      XSSFWorkbook workbook = new XSSFWorkbook(file);  //create a new workbook from the file selected
      XSSFSheet sheet = workbook.getSheetAt(0); // get the first sheet
      Row row;
      processInformationArea.setText(processInformationArea.getText() + "\nReading rows.........");
      for (int i = 0; i <= sheet.getLastRowNum(); i++) {
        row = sheet.getRow(i);
        String email = row.getCell(0).getStringCellValue();
        String withCourse = row.getCell(1).getStringCellValue();
        allUsers.add(email);
        if(!withCourse.equalsIgnoreCase("null")){
          usersWithCourse.add(withCourse);
        }
      }
      processInformationArea.setText(processInformationArea.getText() + "\nReading file  completed.........");
      processInformationArea.setText(processInformationArea.getText() + "\nUsers read........." + allUsers.size());
      processInformationArea.setText(processInformationArea.getText() + "\nUsers with courses read........." + usersWithCourse.size());
      for(int j = 0; j < usersWithCourse.size(); j++){
        if(allUsers.contains(usersWithCourse.get(j))){
          allUsers.remove(usersWithCourse.get(j));
        }
      }
      processInformationArea.setText(processInformationArea.getText() + "\nAfter filtering.........");
      processInformationArea.setText(processInformationArea.getText() + "\nNew Users read........." + allUsers.size());
      //Writing the new File
      XSSFWorkbook xlsWorkbook = new XSSFWorkbook();
      XSSFSheet xlsSheet = xlsWorkbook.createSheet("un-enrolled");
      short rowIndex = 0;
      processInformationArea.setText(processInformationArea.getText() + "\nWriting data to Excel file...");
      for (int i = 0; i < allUsers.size(); i++) {
        XSSFRow dataRow = xlsSheet.createRow(rowIndex++);
        short colIndex = 0;
          dataRow.createCell(colIndex++).setCellValue(allUsers.get(i).toString());
      }
      xlsWorkbook.write(new FileOutputStream(System.getProperty("user.home") + "\\Desktop\\filter-test.xlsx"));

    } catch (IOException e) {
      e.printStackTrace();
    }finally {
      processInformationArea.setText(processInformationArea.getText() + "\nFinished writing file");
    }
  }
}
