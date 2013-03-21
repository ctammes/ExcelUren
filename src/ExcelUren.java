import javax.swing.*;
import java.util.logging.Logger;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 12:13
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUren {

    // initialiseer logger
    public static Logger log = Logger.getLogger(ExcelUren.class.getName());

    public static void main(String[] args) {
        String logDir = ".";
        String logNaam = "ExcelUren.log";

//        try {
//            MijnLog mijnlog = new MijnLog(logDir, logNaam, true);
//            log = mijnlog.getLog();
//            log.setLevel(Level.INFO);
//        } catch (Exception e) {
//            System.out.println(e.getMessage());
//        }

        JFrame frame = new JFrame("ExcelUrenView");
        frame.setContentPane(new ExcelUrenView().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.pack();
        frame.setVisible(true);
    }


}
