//package org.example;
//
//import org.jodconverter.core.DocumentConverter;
//import org.jodconverter.core.office.OfficeException;
//import org.jodconverter.local.LocalConverter;
//import org.jodconverter.local.office.LocalOfficeManager;
//
//import java.io.File;
//
//public class Converter {
//    public static void main(String[] args) {
//        // Укажите путь к установке LibreOffice
//        String officeHome = "C:/Program Files/LibreOffice"; // Измените этот путь в зависимости от вашей системы
//
//        // Создаем экземпляр менеджера офиса
//        LocalOfficeManager officeManager = LocalOfficeManager.builder()
//                .officeHome(new File(officeHome))
//                .install().build();
//
//        try {
//            // Запускаем менеджер офиса
//            officeManager.start();
//
//            // Создаем экземпляр конвертера
//            DocumentConverter converter = LocalConverter.make(officeManager);
//
//            // Указываем входной DOCX файл и выходной PDF файл
//            File inputFile = new File("./files/Maples+Group+-+Transfer+Agent+FAQ+-+August+2023.docx");
//            File outputFile = new File("./output.pdf");
//
//            // Выполняем конвертацию
//            converter.convert(inputFile).to(outputFile).execute();
//
//            System.out.println("Конвертация успешно завершена!");
//        } catch (OfficeException e) {
//            e.printStackTrace();
//        } finally {
//            // Останавливаем менеджер офиса
//            try {
//                officeManager.stop();
//            } catch (OfficeException e) {
//                e.printStackTrace();
//            }
//        }
//    }
//}
