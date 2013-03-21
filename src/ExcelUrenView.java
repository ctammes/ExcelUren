import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FilenameFilter;
import java.util.List;
import java.util.regex.Pattern;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 21-3-13
 * Time: 8:59
 * To change this template use File | Settings | File Templates.
 */
public class ExcelUrenView {
    protected JPanel mainPanel;

    private JTextField txtExcelDir;
    private JCheckBox chkVerlof;
    private JCheckBox chkProject;
    private JComboBox cmbProjecten;
    private JButton btnStart;
    private JList lstResultaat;
    private JButton btnInlezen;

    private ExcelUren uren;
    private Excel excel;

    private String dirXls = "/home/chris/IdeaProjects/uren2012";

    public ExcelUrenView() {

        txtExcelDir.setText(dirXls);

        btnInlezen.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                String xlsDir = txtExcelDir.getText();
                String[] files = leesXlsNamen(xlsDir);
                if (files.length > 0) {
                    excel = new Excel(xlsDir + "/" + files[0]);
                    List<String> projecten = excel.leesProjecten();
                    if (projecten.size() > 0) {
                        cmbProjecten.removeAll();
//                        Arrays.sort(projecten);
                        for (String project: projecten) {
                            cmbProjecten.addItem(project);
                        }
                    }
                }
            }
        });
    }

//    private void createUIComponents() {
//        // TODO: place custom component creation code here
//    }

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


}
