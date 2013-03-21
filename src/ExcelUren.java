import javax.swing.*;
import java.io.File;
import java.io.FilenameFilter;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUren {

    private String dirXls = "../../uren2012";

    private Excel excel;

    public String[] leesXlsNamen(String dirXls) {
        File map = new File(dirXls);
        String[] files = map.list(new FilenameFilter() {
            @Override
            public boolean accept(File map, String fileName) {
                return Pattern.matches("cts\\d+\\.xls", fileName.toLowerCase());
            }
        });
        return files;
    }

    public static void main(String[] args) {
        JFrame frame = new JFrame("ExcelUrenView");
        frame.setContentPane(new ExcelUrenView().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }



}
