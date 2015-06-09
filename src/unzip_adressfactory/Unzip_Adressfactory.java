/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package unzip_adressfactory;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 *
 * @author jungoliver
 */
public class Unzip_Adressfactory {

    /**
     * @param args the command line arguments
     */

 
    /**
     * @param args
     */
    public static void main(String[] args) throws Exception {
        new Unzip_Adressfactory().extractArchive(new String(
                "\\\\Svl-fil02\\msg$\\Msg allgemein\\Vertrieb\\temp_daten"), new File(
                "\\\\Svl-fil02\\msg$\\Msg allgemein\\Vertrieb\\temp_daten"));
    }
 
    public void extractArchive(String archive, File destDir) throws Exception {
        if (!destDir.exists()) {
            destDir.mkdir();
        }
        
        // Das TEmp-Verzeichnis wird durchsucht, eine ZIP-Datei geöffnet wenn gefunden und die darin befindliche
        // XLS-Datei entpackt, wenn vorhanden
        File[] ListFiles= new File(archive).listFiles();
        for(int i=0;i<ListFiles.length;i++)
        {
         if(ListFiles[i].getName().contains((".zip")))
         {
        
        ZipFile zipFile = new ZipFile(ListFiles[i]);
        Enumeration entries = zipFile.entries();
 
        byte[] buffer = new byte[16384];
        int len;
        while (entries.hasMoreElements()) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
 
            String entryFileName = entry.getName();
            if (entryFileName.contains("xls"))
            {
            File dir = dir = buildDirectoryHierarchyFor(entryFileName, destDir);
            if (!dir.exists()) {
                dir.mkdirs();
            }
 
            if (!entry.isDirectory()) {
                BufferedOutputStream bos = new BufferedOutputStream(
                        new FileOutputStream(new File(destDir, entryFileName)));
 
                BufferedInputStream bis = new BufferedInputStream(zipFile
                        .getInputStream(entry));
 
                while ((len = bis.read(buffer)) > 0) {
                    bos.write(buffer, 0, len);
                }
 
                bos.flush();
                bos.close();
                bis.close();
            }
            }
        }
                zipFile.close();
                ListFiles[i].delete();
               
         }
        }
        
        //Die gefundene XLS-Datei wird, wenn sie auf _1_Abos_fertig_ endet, geöffnet und
        //in "1_Abos_fertig_.xls umbenannt
        File[] ListFiles2= new File(archive).listFiles();
        for(int i=0;i<ListFiles2.length;i++)
        {
         if(ListFiles2[i].getName().contains(("_1_Abos_fertig_.xls")))
         {
             ListFiles2[i].renameTo(new File(archive+"\\1_Abos_fertig_.xls"));
         }
        }
    }
 
    private File buildDirectoryHierarchyFor(String entryName, File destDir) {
        int lastIndex = entryName.lastIndexOf('/');
        String entryFileName = entryName.substring(lastIndex + 1);
        String internalPathToEntry = entryName.substring(0, lastIndex + 1);
        return new File(destDir, internalPathToEntry);
    }
}
    
    

