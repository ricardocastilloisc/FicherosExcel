import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Scanner;
public class Comparaciones {

	
	static String [][] basededatosRecorrido;	
	
	static String [][] basededatosComparar;	
	
	static String [][] basededatosTerminada;
	
	
	
	static int tamanoDeFilasDeBaseDeDatosProgramadas = 0;
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		
        String fileName1= "basedeDatosSinRfc.xls";
        String fileName2 = "baseDeDatosConRFC.xls";
       
        
        //CreateExcel(fileName);
        
        ReadExcel(fileName1, 1);
        ReadExcel(fileName2, 2);
        
        
        basededatosTerminada = new String[basededatosRecorrido.length][ basededatosComparar[0].length + 1];
        
        
        acomodoDeArray(); 
        
        
        CreateExcel("BaseDeDatosTerminada.xls");
        
        
        /*
        for(int recorrido = 0; recorrido<basededatosTerminada.length; recorrido++) 
        {
        	
        	for(int recorridoInterno = 0; recorridoInterno < basededatosTerminada[0].length; recorridoInterno++) 
        	{
        		System.out.print( "||" + basededatosTerminada[recorrido][recorridoInterno] );
        	}
        	System.out.println();
        }
        */
        //OverwriteExcel(fileName, data);
        
        //AcomodarArreglo();
        
        //CreateExcel("BaseDeDatosCapacitacionProcesada.xls");
         
	}
	
	  public static void acomodoDeArray() 
      {
		  for(int recorrido = 0; recorrido < basededatosRecorrido.length; recorrido++ ) 
		  {
			  String string =  basededatosRecorrido[recorrido][0];
			  //String[] parts = string.split(" ");
			  
			  String PalabraAcomparar  =  basededatosRecorrido[recorrido][0];
			  
			  
			 // System.out.println(PalabraAcomparar);
			  /*
			  if(parts.length < 5) 
			  {
				   if(parts.length == 4) 
				   {
					   PalabraAcomparar =  parts[2] + " " +parts[3] + " " + parts[0] + " " + parts[1];
				  }
			
				  if(parts.length == 3) 
				  {
					  PalabraAcomparar =  parts[1] + " " + parts[2] + " " + parts[0];
				  }
				  
				  if(parts.length == 2) 
				  {
					  PalabraAcomparar =  parts[1] + " " + parts[0];
				  }
				  
				  basededatosTerminada[recorrido][0] = PalabraAcomparar;
			  }else 
			  {
				  basededatosTerminada[recorrido][0] = basededatosRecorrido[recorrido][0];
			  }
			  
			  */
			  
			  basededatosTerminada[recorrido][0] = PalabraAcomparar;
			 
			 
			  
			  for(int recorridoInternoExterior  = 0; recorridoInternoExterior < basededatosComparar.length; recorridoInternoExterior++) 
			  {
				   if
				   (
						   PalabraAcomparar.equals(basededatosComparar[recorridoInternoExterior][5]) ||
						   PalabraAcomparar.equals(basededatosComparar[recorridoInternoExterior][5] + " ") ||
						   PalabraAcomparar.equals(basededatosComparar[recorridoInternoExterior][5] + "  ") ||
						   PalabraAcomparar.equals(basededatosComparar[recorridoInternoExterior][5] + "   ")
					) 
				   {
					   for(int recorridoInterno = 0; recorridoInterno < basededatosComparar[0].length; recorridoInterno++) 
					   {
						   basededatosTerminada[recorrido][recorridoInterno+1] = basededatosComparar[recorridoInternoExterior][recorridoInterno];
					   }
				   }
			  }
		  }
      }
	  
	  public static String method(String str) {
	    if (str != null && str.length() > 0 && str.charAt(str.length()-1)==',') {
	      str = str.substring(0, str.length()-1);
	    }
	    return str;
	}

    private static void ReadExcel(String fileName, int numeroBaseDeDatos) {
        try {
            InputStream myFile = new FileInputStream(new File(fileName));
            HSSFWorkbook wb = new HSSFWorkbook(myFile);
            HSSFSheet sheet = wb.getSheetAt(0);

            HSSFCell cell;
            HSSFRow row;
            
            
            //filas luego //columnas
            
            if(numeroBaseDeDatos == 1) 
            
            {
                
            	basededatosRecorrido = new String[sheet.getLastRowNum() + 1][sheet.getRow(0).getLastCellNum()];
                
                
                tamanoDeFilasDeBaseDeDatosProgramadas =  sheet.getLastRowNum() + 1;
                //System.out.println("Apunto de entrar a loops");

                //System.out.println("" + sheet.getLastRowNum());

                for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                    row = sheet.getRow(i);
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        cell = row.getCell(j);
                        basededatosRecorrido[i][j] = cell.toString();
                        //System.out.print("|" +  basededatos[i][j] + "|");
                    }
                    //System.out.println("");
                }
                System.out.println("Finalizado");
            }else 
            {
            	
            	basededatosComparar = new String[sheet.getLastRowNum() + 1][sheet.getRow(0).getLastCellNum()];
                tamanoDeFilasDeBaseDeDatosProgramadas =  sheet.getLastRowNum() + 1;
                //System.out.println("Apunto de entrar a loops");

                //System.out.println("" + sheet.getLastRowNum());

                for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
                    row = sheet.getRow(i);
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        cell = row.getCell(j);
                        basededatosComparar[i][j] = cell.toString();
                        //System.out.print("|" +  basededatos[i][j] + "|");
                    }
                    //System.out.println("");
                }
                System.out.println("Finalizado");
            }
            
       

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
    
    
    private static void CreateExcel(String fileName) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("Cursos");
        
        String[][] data = {{"1","2"},{"3","4"}};
        
        
        //System.out.println(data.length);

        
      for (int j = 0; j < basededatosTerminada.length; j++)
        
      {// 2 por el Encabezado y la linea de informacion
            
    	
    	  HSSFRow row = sheet.createRow(j);
 
            for (int i = 0; i < basededatosTerminada[0].length; i++) 
            
            {// Tantos loops como info en el arreglo
            
            	
            	HSSFCell cell = row.createCell(i);
            
                cell.setCellValue(basededatosTerminada[j][i]);
                
      
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


}
