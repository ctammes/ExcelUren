import javax.swing.*;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 21-3-13
 * Time: 20:56
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUrenMain {

    public static void main(String[] args) {

        JFrame frame = new JFrame("ExcelUrenView");
        frame.setContentPane(new ExcelUrenView().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        frame.setSize();
        frame.pack();
        frame.setVisible(true);

    }

}
