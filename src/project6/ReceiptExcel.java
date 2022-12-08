package project6;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReceiptExcel extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread { // Поток запуска MS Excel

        public void run() {
            
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator"); // Текущий катаолог
            try {
                modifData(dir + "receipt_template.xls", dir + "receipt.xls", 
                        jTextField_FIO.getText(),
                        jTextField_Vacancy.getText(), 
                        jTextField_Salary1.getText(), 
                        jTextField_Employment.getText(),
                        jTextField_Adres.getText(),
                        jTextField_Number.getText(),
                        jTextField_Mail.getText(),
                        jTextField_Citizenship.getText(),
                        jTextField_Education.getText(),
                        jTextField_Data.getText(),
                        jTextField_Status.getText(),
                        jTextField_Year.getText(),
                        jTextField_Place.getText(),
                        jTextField_Faculty.getText(),
                        jTextField_Specialization.getText(),
                        jTextField_Gender.getText()); // Вызов метода создания отчета
                if (System.getProperty("os.name").equals("Linux")
                        && System.getProperty("java.vendor").startsWith("Red Hat")) {
                    new ProcessBuilder("xdg-open", dir + "receipt.xls").start();
                } else {
                    Desktop.getDesktop().open(new File(dir + "receipt.xls")); // Запуск отчета в MS Excel
                }
            } catch (Exception ex) {
                System.err.println("Error modifData!");
                ex.printStackTrace();
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    // Метод создания отчета
    private void modifData(String inputFileName, String outputFileName, String FIO, String vacancy,
            String salary, String employment, String adres, String number,
            String mail, String citizenship, String education, String data,
            String status, String year, String place, String faculty, String specialization, String gender) throws IOException {
        HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(new FileInputStream(inputFileName))); // Документ MS Excel
        HSSFSheet sheet = wb.getSheetAt(0); // Первый лист в документе MS Excel
        sheet.getRow(0).getCell(3).setCellValue(FIO);
        sheet.getRow(3).getCell(3).setCellValue(vacancy);
        sheet.getRow(5).getCell(2).setCellValue(salary);
        sheet.getRow(7).getCell(2).setCellValue(employment);
        sheet.getRow(9).getCell(2).setCellValue(number);
        sheet.getRow(11).getCell(2).setCellValue(mail);
        sheet.getRow(14).getCell(2).setCellValue(citizenship);
        sheet.getRow(16).getCell(2).setCellValue(adres);
        sheet.getRow(18).getCell(2).setCellValue(education);
        sheet.getRow(20).getCell(2).setCellValue(data);
        sheet.getRow(22).getCell(2).setCellValue(gender);
        sheet.getRow(24).getCell(2).setCellValue(status);
        sheet.getRow(26).getCell(2).setCellValue(year);
        sheet.getRow(28).getCell(2).setCellValue(place);
        sheet.getRow(30).getCell(2).setCellValue(faculty);
        sheet.getRow(32).getCell(2).setCellValue(specialization);
        
        try (FileOutputStream fileOut = new FileOutputStream(outputFileName)) {
            wb.write(fileOut);
        }
    }

    public ReceiptExcel() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jButton1 = new javax.swing.JButton();
        jTextField_FIO = new javax.swing.JTextField();
        jTextField_Vacancy = new javax.swing.JTextField();
        jTextField_Salary1 = new javax.swing.JTextField();
        jTextField_Mail = new javax.swing.JTextField();
        jTextField_Adres = new javax.swing.JTextField();
        jTextField_Employment = new javax.swing.JTextField();
        jTextField_Number = new javax.swing.JTextField();
        jTextField_Citizenship = new javax.swing.JTextField();
        jTextField_Education = new javax.swing.JTextField();
        jTextField_Data = new javax.swing.JTextField();
        jTextField_Status = new javax.swing.JTextField();
        jTextField_Year = new javax.swing.JTextField();
        jTextField_Place = new javax.swing.JTextField();
        jTextField_Faculty = new javax.swing.JTextField();
        jTextField_Specialization = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        jLabel4 = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        jTextField_Gender = new javax.swing.JTextField();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jLabel14 = new javax.swing.JLabel();
        jLabel15 = new javax.swing.JLabel();
        jLabel16 = new javax.swing.JLabel();
        jLabel17 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Работа с Excel");
        setResizable(false);
        getContentPane().setLayout(null);

        jButton1.setText("в Excel");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        getContentPane().add(jButton1);
        jButton1.setBounds(460, 530, 72, 22);

        jTextField_FIO.setFont(new java.awt.Font("Tahoma", 0, 12)); // NOI18N
        getContentPane().add(jTextField_FIO);
        jTextField_FIO.setBounds(230, 40, 180, 21);

        jTextField_Vacancy.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Vacancy);
        jTextField_Vacancy.setBounds(230, 70, 180, 20);

        jTextField_Salary1.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Salary1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_Salary1ActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Salary1);
        jTextField_Salary1.setBounds(300, 100, 140, 20);

        jTextField_Mail.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Mail);
        jTextField_Mail.setBounds(300, 190, 140, 20);

        jTextField_Adres.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Adres);
        jTextField_Adres.setBounds(180, 300, 140, 20);

        jTextField_Employment.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Employment);
        jTextField_Employment.setBounds(300, 130, 140, 20);

        jTextField_Number.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Number);
        jTextField_Number.setBounds(300, 160, 140, 20);

        jTextField_Citizenship.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Citizenship);
        jTextField_Citizenship.setBounds(180, 270, 140, 20);

        jTextField_Education.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Education);
        jTextField_Education.setBounds(180, 330, 140, 20);

        jTextField_Data.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Data);
        jTextField_Data.setBounds(180, 360, 140, 20);

        jTextField_Status.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Status);
        jTextField_Status.setBounds(180, 420, 140, 20);

        jTextField_Year.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        jTextField_Year.setToolTipText("");
        getContentPane().add(jTextField_Year);
        jTextField_Year.setBounds(180, 450, 140, 20);

        jTextField_Place.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Place);
        jTextField_Place.setBounds(180, 480, 140, 20);

        jTextField_Faculty.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Faculty);
        jTextField_Faculty.setBounds(180, 510, 140, 20);

        jTextField_Specialization.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Specialization);
        jTextField_Specialization.setBounds(180, 540, 140, 20);

        jLabel2.setText("ФИО");
        getContentPane().add(jLabel2);
        jLabel2.setBounds(150, 40, 80, 20);

        jLabel3.setText("Занятость");
        getContentPane().add(jLabel3);
        jLabel3.setBounds(210, 100, 90, 20);

        jLabel4.setText("Вакансия");
        getContentPane().add(jLabel4);
        jLabel4.setBounds(150, 70, 80, 20);

        jLabel5.setText("Специальность");
        getContentPane().add(jLabel5);
        jLabel5.setBounds(50, 540, 130, 20);

        jLabel6.setText("График");
        getContentPane().add(jLabel6);
        jLabel6.setBounds(210, 130, 90, 20);

        jLabel7.setText("Номер");
        getContentPane().add(jLabel7);
        jLabel7.setBounds(210, 160, 90, 20);

        jLabel8.setText("Почта");
        getContentPane().add(jLabel8);
        jLabel8.setBounds(210, 190, 90, 20);

        jLabel9.setText("Гражданство");
        getContentPane().add(jLabel9);
        jLabel9.setBounds(50, 270, 130, 20);

        jLabel10.setText("Город");
        getContentPane().add(jLabel10);
        jLabel10.setBounds(50, 300, 130, 20);

        jLabel11.setText("Образование");
        getContentPane().add(jLabel11);
        jLabel11.setBounds(50, 330, 130, 20);

        jTextField_Gender.setFont(new java.awt.Font("Tahoma", 0, 14)); // NOI18N
        getContentPane().add(jTextField_Gender);
        jTextField_Gender.setBounds(180, 390, 140, 20);

        jLabel12.setText("Дата рождения");
        getContentPane().add(jLabel12);
        jLabel12.setBounds(50, 360, 130, 20);

        jLabel13.setText("Пол");
        getContentPane().add(jLabel13);
        jLabel13.setBounds(50, 390, 130, 20);

        jLabel14.setText("Семейное положение");
        getContentPane().add(jLabel14);
        jLabel14.setBounds(50, 420, 130, 20);

        jLabel15.setText("Год выпуска");
        getContentPane().add(jLabel15);
        jLabel15.setBounds(50, 450, 130, 20);

        jLabel16.setText("ВУЗ");
        getContentPane().add(jLabel16);
        jLabel16.setBounds(50, 480, 130, 20);

        jLabel17.setText("Факультет");
        getContentPane().add(jLabel17);
        jLabel17.setBounds(50, 510, 130, 20);

        setSize(new java.awt.Dimension(648, 651));
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton1ActionPerformed

    private void jTextField_Salary1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_Salary1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_Salary1ActionPerformed

    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Windows".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ReceiptExcel.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        
        
        //</editor-fold>
        

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ReceiptExcel().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JTextField jTextField_Adres;
    private javax.swing.JTextField jTextField_Citizenship;
    private javax.swing.JTextField jTextField_Data;
    private javax.swing.JTextField jTextField_Education;
    private javax.swing.JTextField jTextField_Employment;
    private javax.swing.JTextField jTextField_FIO;
    private javax.swing.JTextField jTextField_Faculty;
    private javax.swing.JTextField jTextField_Gender;
    private javax.swing.JTextField jTextField_Mail;
    private javax.swing.JTextField jTextField_Number;
    private javax.swing.JTextField jTextField_Place;
    private javax.swing.JTextField jTextField_Salary1;
    private javax.swing.JTextField jTextField_Specialization;
    private javax.swing.JTextField jTextField_Status;
    private javax.swing.JTextField jTextField_Vacancy;
    private javax.swing.JTextField jTextField_Year;
    // End of variables declaration//GEN-END:variables
}
