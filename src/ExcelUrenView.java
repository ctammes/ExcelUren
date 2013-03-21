import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FilenameFilter;
import java.util.*;
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
    private JCheckBox chkTijdInUit;

    private ExcelUren uren;
    private Excel excel;

    private String dirXls = "/home/chris/IdeaProjects/uren2012";
    private String[] files = null;

    public ExcelUrenView() {

        txtExcelDir.setText(dirXls);

        btnInlezen.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                // lees en sorteer alle xls-bestanden
                files = leesXlsNamen(txtExcelDir.getText());
                Arrays.sort(files);

                if (files.length > 0) {
                    // per bestand projecten inlezen in gesorteerde lijst (TreeSet) zonder duplicaten (Set)
                    Set<String> projecten = new TreeSet<String>();
                    for (String xlsFile: files) {
                        excel = new Excel(dirXls + "/" + xlsFile);
                        // nieuwe projecten toevoegen - geen duplicaten (want Set)
                        projecten.addAll(excel.leesProjecten());
                        excel.sluitWerkblad();
                    }
                    if (projecten.size() > 0) {
                        cmbProjecten.removeAll();
                        for (String project: projecten) {
                            //TODO wat is een ComboBoxModel ??
                            cmbProjecten.addItem(project);
                        }
                    }
                }
            }
        });
        btnStart.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                String project = "";
                if (chkVerlof.isSelected()) {
                    project = "verlof";
                } else if (chkProject.isSelected() && cmbProjecten.getSelectedIndex()>=0) {
                    project = cmbProjecten.getSelectedItem().toString();

                }

                // verwerk de files
                int granttotal = 0;
                DefaultListModel listModel = new DefaultListModel();
                for (String xlsFile: files) {
                    excel = new Excel(dirXls + "/" + xlsFile);
                    int totaal = excel.totaliseerDuur(project);
                    float dagtotaal = excel.totaliseerDagtotaal() * 24;
                    if (project == "verlof" || (project != "verlof" && totaal > 0)) {
                        granttotal += totaal;
                        String tekst = String.format("file: %s, minuten: %4d, uren: %2.1f, dagen: %2.2f, uren gewerkt: %2.0f \n", xlsFile, totaal, (float) totaal / 60, (float) totaal / 60 / 9, dagtotaal);
                        listModel.addElement(tekst);
                    }
                    excel.sluitWerkblad();
                }
                String tekst = String.format("Totaal: uren: %d, dagen: %d", granttotal / 60, granttotal / 60 / 9);
                listModel.addElement(tekst);

                lstResultaat.setModel(listModel);
            }
        });
        chkProject.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                cmbProjecten.setEnabled(chkProject.isEnabled());
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
