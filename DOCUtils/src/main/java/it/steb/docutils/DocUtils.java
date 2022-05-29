/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Project/Maven2/JavaApp/src/main/java/${packagePath}/${mainClassName}.java to edit this template
 */

package it.steb.docutils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;


/**
 *
 * @author stefa
 */
public class DocUtils {
    
    
    /**
     * Converts .docx file to PDF
     * @param input .docx File
     * @param out .pdf where the file will be created
     * @return returns the PDF file
     * @throws Exception
     */
    public File DOCUConvertToPDF(File input, File out) throws Exception {
        
        InputStream inputStream = new FileInputStream(input);
        WordprocessingMLPackage word = WordprocessingMLPackage.load(inputStream);
        
        
        FileOutputStream os = new FileOutputStream(out);
        Docx4J.toPDF(word, os);
        os.flush();
        os.close();
        
        return out;
    }
    
    
    /**
     * Takes word file and modifies variables into it
     * @param input Word input file
     * @param out Word modified output file
     * @param vars Variables to be replaced inside the word file
     * @return Word modified file
     * @throws Exception 
     */
    public File DOCUReplacer(File input, File out, HashMap vars) throws Exception {
        
        WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(new FileInputStream(input));
        wordMLPackage.getMainDocumentPart().variableReplace(vars);
        if(!out.exists()) out.createNewFile();
            
            
        wordMLPackage.save(out);
        
        return out;
        
    }
    
}
