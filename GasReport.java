package gasreport;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;

public class GasReport {

	public static void main(String[] args) throws java.io.IOException {
        	JFileChooser jfc = new JFileChooser("C:\\");
        	jfc.setDialogTitle("Gas Report: Browse Files");
        	jfc.setAcceptAllFileFilterUsed(false);
       		FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Spreadsheets", "xlsx");
        	jfc.addChoosableFileFilter(filter);
        
        	int returnValue = jfc.showDialog(null, "Submit");
        	if (returnValue == JFileChooser.APPROVE_OPTION) {
	        	File selectedFile = jfc.getSelectedFile();
			String filepath = selectedFile.getAbsolutePath();

            		ProcessBuilder pb = new ProcessBuilder("python", "C:\\Python36\\gasreport.py", filepath);
            		Process p = pb.start();
		}
	}
}  
