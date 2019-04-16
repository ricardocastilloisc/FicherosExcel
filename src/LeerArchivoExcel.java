
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Scanner;


/**
 * Created by Leon on 07/02/2017.
 */


public class LeerArchivoExcel {

	static String [][] basededatos;	
	
	static String [][] basededatosProcesada;	
	
	
	static int tamanoDeFilasDeBaseDeDatosProgramadas = 0;
	
	
	public static void main(String[] args) {
        String fileName = "basedeDatosCapacitacion.xls";
       
        
        //CreateExcel(fileName);
        ReadExcel(fileName);
        //OverwriteExcel(fileName, data);
        
        AcomodarArreglo();
        
        //CreateExcel("BaseDeDatosCapacitacionProcesada.xls");
         
        
    }
	
	
	private static void AcomodarArreglo() 
	{
		
		String year2017 =  "";
		String year2018 =  "";
		
		String erroresNo = "";
		
		basededatosProcesada = new String [tamanoDeFilasDeBaseDeDatosProgramadas][3];
		
		for(int recorridoFilas = 0; recorridoFilas < tamanoDeFilasDeBaseDeDatosProgramadas; recorridoFilas++)
		{
			
			basededatosProcesada[recorridoFilas][0] = basededatos[recorridoFilas][0];
			
			
			int recorridoColumnasW = 1;
			
			while(recorridoColumnasW < basededatos[0].length) 
			{
				if(
						basededatos[recorridoFilas][recorridoColumnasW].equals("OK") ||
						basededatos[recorridoFilas][recorridoColumnasW].equals("OK ") ||
						basededatos[recorridoFilas][recorridoColumnasW].equals(" OK") ||
						basededatos[recorridoFilas][recorridoColumnasW].equals("OK  ") ||
						basededatos[recorridoFilas][recorridoColumnasW].equals("  OK")
						) 
				{
					if (basededatos[recorridoFilas][recorridoColumnasW + 1].equals("2017.0")) 
					{
						year2017 = year2017 +basededatos[0][recorridoColumnasW] + ", ";
					}else{
						if (basededatos[recorridoFilas][recorridoColumnasW + 1].equals("no")) 
						{
							erroresNo = erroresNo + basededatos[recorridoFilas][0] + ", ";
						}else 
						{
							year2018 = year2018 + basededatos[0][recorridoColumnasW] + ", ";
						}
						
					}
				}
				
				recorridoColumnasW = recorridoColumnasW + 2;
			}
			
			/*
			for (int recorridoColumnas = 1; recorridoColumnas < basededatos[0].length; recorridoColumnas++) 
			{
				if(basededatos[recorridoFilas][recorridoColumnas].equals("OK")) 
				{
					if (basededatos[recorridoFilas][recorridoColumnas + 1].equals("2017.0")) 
					{
						year2017 = basededatos[0][recorridoColumnas] + "/";
					}else{
						if (basededatos[recorridoFilas][recorridoColumnas + 1].equals("no")) 
						{
							erroresNo = basededatos[0][0] + "/";
						}else 
						{
							year2018 = basededatos[0][recorridoColumnas] + "/";
						}
						
					}
				}
				recorridoColumnas++;
			}
			*/
			
			if(!year2017.equals("")) 
			{
				year2017 = year2017.substring(0, year2017.length()-2);
			}
			
			if(!year2018.equals("")) 
			{
				year2018 = year2018.substring(0, year2018.length()-2);
			}
			
			
			basededatosProcesada[recorridoFilas][1] = year2017;
			
			
			basededatosProcesada[recorridoFilas][2] = year2018;
			
			
			year2017 = "";
			year2018 = "";
		}
		
		
		/* for(int recorridoFilas2 = 0; recorridoFilas2<basededatosProcesada.length; recorridoFilas2++ ) 
		 {
			 for(int recorridoColumnas2 = 0; recorridoColumnas2<3; recorridoColumnas2++) 
			 {
				 System.out.print("[" + basededatosProcesada[recorridoFilas2][recorridoColumnas2] + " ] ");
			 }
			 
			 System.out.println("");
		 }
		*/
		
		System.out.println("termino");
	}
	
	public static String method(String str) {
	    if (str != null && str.length() > 0 && str.charAt(str.length()-1)==',') {
	      str = str.substring(0, str.length()-1);
	    }
	    return str;
	}

    private static void CreateExcel(String fileName) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Cursos");
        
        String[][] data = {{"1","2"},{"3","4"}};
        
        
        //System.out.println(data.length);

        
      for (int j = 0; j < basededatosProcesada.length; j++)
        
      {// 2 por el Encabezado y la linea de informacion
            
    	
    	  HSSFRow row = sheet.createRow(j);
 
            for (int i = 0; i < basededatosProcesada[0].length; i++) 
            
            {// Tantos loops como info en el arreglo
            
            	
            	HSSFCell cell = row.createCell(i);
            
                cell.setCellValue(basededatosProcesada[j][i]);
                
      
            }
        
      }

        try {
            FileOutputStream fos = null;
            File file;

            file = new File(fileName);
            fos = new FileOutputStream(file);

            workbook.write(fos);
            workbook.close();
            fos.close();
            System.out.println("Finalizado");

        } catch (Exception e) {
            // TODO: handle exception
            System.out.println(e.getMessage());
        }
    }

    private static void ReadExcel(String fileName) {
        try {
            InputStream myFile = new FileInputStream(new File(fileName));
            HSSFWorkbook wb = new HSSFWorkbook(myFile);
            HSSFSheet sheet = wb.getSheetAt(0);

            HSSFCell cell;
            HSSFRow row;
            
            
            //filas luego //columnas
            
            
            basededatos = new String[sheet.getLastRowNum() + 1][sheet.getRow(0).getLastCellNum()];
            
            
            tamanoDeFilasDeBaseDeDatosProgramadas =  sheet.getLastRowNum() + 1;
            //System.out.println("Apunto de entrar a loops");

            //System.out.println("" + sheet.getLastRowNum());

            for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                row = sheet.getRow(i);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    cell = row.getCell(j);
                    basededatos[i][j] = cell.toString();
                    //System.out.print("|" +  basededatos[i][j] + "|");
                }
                //System.out.println("");
            }
            System.out.println("Finalizado");

        } catch (Exception e) {
            // TODO: handle exception
            System.out.println(e.getMessage());
        }
    }

    private static void OverwriteExcel(String fileName, String[] data) {
        try {
            InputStream inp = new FileInputStream(new File(fileName));
            HSSFWorkbook oldWorkbook = new HSSFWorkbook(inp);

            HSSFSheet oldSheet = oldWorkbook.getSheetAt(0);

            HSSFRow oldRow;

            oldRow = oldSheet.createRow(oldSheet.getLastRowNum() + 1);
            for (int i = 0; i < data.length; i++) {// Tantos loops como info en el arreglo
                HSSFCell cell = oldRow.createCell(i);
                cell.setCellValue(data[i]);
            }

            FileOutputStream fos = null;
            File file;

            file = new File(fileName);
            fos = new FileOutputStream(file);

            oldWorkbook.write(fos);
            oldWorkbook.close();
            fos.close();

            System.out.println("Finalizado");

        } catch (Exception e) {
            // TODO: handle exception
            System.out.println(e.getMessage());
        }
    }
}