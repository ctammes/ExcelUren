import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;

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

    public ExcelUrenView() {
        btnInlezen.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                uren = new ExcelUren();

                String xlsDir = txtExcelDir.toString();
                String[] files = uren.leesXlsNamen(xlsDir);
                if (files.length > 0) {
                    excel = new Excel(xlsDir + "/ " + files[0]);
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

    private void createUIComponents() {
        // TODO: place custom component creation code here
    }
}
