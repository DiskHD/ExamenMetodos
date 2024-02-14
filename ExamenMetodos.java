/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package examenmetodos;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Scanner;
import java.util.ArrayList;
import org.apache.poi.xssf.usermodel.*;
import java.io.*;
import java.awt.Desktop;
import java.io.File;
import java.io.IOException;

import java.io.FileOutputStream;


import java.io.FileOutputStream;
import java.io.IOException;
/**
 *
 * @author hdzdi
 */
public class ExamenMetodos {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        Scanner sc=new Scanner(System.in);
        ArrayList<Integer>vector=new ArrayList<>();
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("MiHoja");
        System.out.println("Dame el número de datos");
        double[]datos=new double[sc.nextInt()];
        //rellenar los datos
        for(int i=0;i<datos.length;i++)
            datos[i]=sc.nextDouble();
        
        
        
        
        //se rellenan la primera fila
        Row row=sheet.createRow(0);
        Cell cell;
        double promedio=0;
        for (int i = 0; i < datos.length; i++) {
            promedio+=datos[i];
        }
        promedio=promedio/datos.length;
        System.out.println("Promedio: "+promedio);
        //
        System.out.println("Dime que quieres:");
        boolean b=true;
        while(b==true){
            System.out.println("1. Valor Verdadero(promedio)");
            System.out.println("2. Error");
            System.out.println("3. Error Absoluto");
            System.out.println("4. Error Relativo");
            System.out.println("5. Error Relativo con porcentaje");
            System.out.println("6. Salir");
            int m=sc.nextInt();
            switch(m){
                case 1:
                    vector.add(1);
                    break;
                case 2:
                    vector.add(2);
                    break;
                case 3:
                    vector.add(3);
                    break;
                case 4:
                    vector.add(4);
                    break;
                case 5:
                    vector.add(5);
                    break;
                case 6:
                    b=false;
                    break;
                default:
                    System.out.println("Error");
                    break;
            }
        }
        /*double[][]tabla=new double[vector.size()][datos.length];
        //primeros datos
        for(int i=0;i<datos.length;i++){
            tabla[0][i]=datos[i];
        }*/
        //
       
        
        //
        //////cell = row.createCell(1);
        //row = sheet.createRow(2);//Columna(Vertical)
        //cell = row.createCell(2);//fila(Orizontal)
        //////cell.setCellValue(5);
        //cell2.setCellFormula("AVERAGE(A1:A" + datos.length + ")");    
        //FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        //evaluator.evaluateFormulaCell(cell);
        
        
        
        //Aqui otro for para las columnas
        double Ea=0;
        FormulaEvaluator evaluator ;
        for(int j=0;j<datos.length;j++){
            row = sheet.createRow(j);//Columna(Vertical)
            cell = row.createCell(0);//fila(Orizontal)
            cell.setCellValue(datos[j]);
            for(int i=0;i<vector.size();i++){
                //cell = row.createCell(i);//fila(Orizontal)
                //cell.setCellValue(datos[j]);
                switch(vector.get(i)){
                    case 1:

                        //tabla[i][1]=promedio;
                        cell=row.createCell(i+1);
                        cell.setCellValue(promedio);



                        /*for(int j=0;j<datos.length;j++){
                            cell = row.createCell(i+1);

                            cell.setCellFormula("AVERAGE(A1:A" + datos.length + ")");    
                            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                            evaluator.evaluateFormulaCell(cell);
                            /*
                            row = sheet.createRow(j);//Columna(Vertical)
                            cell = row.createCell(i+1);//fila(Orizontal)
                            cell.setCellValue(5);

                        }
                        */
                        break;
                    case 2://Error
                        cell=row.createCell(i+1);
                        cell.setCellValue(promedio-datos[j]);
                        //cell.setCellFormula("ABS("+columnas(i-1)+"1-B1)");
                        //evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        //evaluator.evaluateFormulaCell(cell);
                        break;
                    case 3://Error Absoluto
                        cell=row.createCell(i+1);
                        Ea=(promedio-datos[j]);
                        if(Ea<0)
                            Ea=Ea*-1;
                        cell.setCellValue(Ea);
                        /*cell.setCellFormula("ABS("+columnas(i-1)+"1-B1)");
                        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        evaluator.evaluateFormulaCell(cell);*/
                        break;
                    case 4:
                        cell=row.createCell(i+1);
                        Ea=(promedio-datos[j]);
                        if(Ea<0)
                            Ea=Ea*-1;
                        cell.setCellValue(Ea/promedio);
                        
                        /*cell.setCellFormula("ABS("+columnas(i-1)+"1-B1)/B1");
                        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        evaluator.evaluateFormulaCell(cell);*/
                        break;
                    case 5:
                        cell=row.createCell(i+1);
                        cell=row.createCell(i+1);
                        Ea=(promedio-datos[j]);
                        if(Ea<0)
                            Ea=Ea*-1;
                        cell.setCellValue((Ea/promedio)*100);
                        
                        /*cell.setCellFormula("ABS(("+columnas(i-1)+"1-B1)/B1)*100");
                        evaluator = workbook.getCreationHelper().createFormulaEvaluator();
                        evaluator.evaluateFormulaCell(cell);*/
                        break;

                }
            }        
        }
       
        // Guardar el libro de trabajo en un archivo
        try (FileOutputStream fileOut = new FileOutputStream(System.getProperty("user.home") + "/OneDrive/Escritorio/DocumentoExcel.xlsx")) {
            workbook.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Cerrar el libro de trabajo
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        System.out.println("Documento Excel generado exitosamente.");
        
        
        // Ruta al archivo que quieres abrir
        String filePath = "C:/Users/hdzdi/OneDrive/Escritorio/DocumentoExcel.xlsx";

        // Crear un objeto File con la ruta del archivo
        File file = new File(filePath);

        // Verificar si el archivo existe
        if (file.exists()) {
            // Obtener el escritorio del sistema
            Desktop desktop = Desktop.getDesktop();

            try {
                // Abrir el archivo con la aplicación predeterminada
                desktop.open(file);
                System.out.println("Archivo abierto exitosamente.");
            } catch (IOException e) {
                System.err.println("Error al abrir el archivo: " + e.getMessage());
            }
        } else {
            System.err.println("El archivo no existe.");
        }
        
        try {
            Thread.sleep(5000); // 5000 milisegundos = 5 segundos
        } catch (InterruptedException e) {
            // Manejar la excepción si ocurre
            e.printStackTrace();
        }
        
        
    }
    public static String columnas(int i){
        switch(i){
            case 0:
                return "A";
            case 1:
                return "B";
            case 2:
                return "C";
            case 3:
                return "D";
            case 4:
                return "E";
            case 5:
                return "F";
            case 6:
                return "G";
            case 7:
                return "H";
            case 8:
                return "I";
            case 9:
                return "J";
            case 10:
                return "K";
            case 11:
                return "L";
            default:
                return "x";
        }
    }
    
}
