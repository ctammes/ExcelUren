import nl.ctammes.common.Diversen;
import nl.ctammes.common.MijnLog;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.Arrays;
import java.util.Set;
import java.util.TreeSet;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 21-3-13
 * Time: 8:59
 * To change this template use File | Settings | File Templates.
 */
public class MijnUren {

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

    private String dirXls = "/home/chris/IdeaProjects/uren2012";
    private String[] files = null;

    public MijnUren() {

        txtExcelDir.setText(dirXls);

        btnInlezen.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent actionEvent) {
                // lees en sorteer alle xls-bestanden
                files = Diversen.leesFileNamen(txtExcelDir.getText(), ExcelUren.URENMASK);
                Arrays.sort(files);

                if (files.length > 0) {
                    // per bestand projecten inlezen in gesorteerde lijst (TreeSet) zonder duplicaten (Set)
                    Set<String> projecten = new TreeSet<String>();
                    for (String xlsFile: files) {
                        uren = new ExcelUren(dirXls, xlsFile);
                        // nieuwe projecten toevoegen - geen duplicaten (want Set)
                        projecten.addAll(uren.leesProjecten());
                        uren.sluitWerkblad();
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
                    uren = new ExcelUren(dirXls, xlsFile);
                    float totaal = uren.geefTaakDuur(project);
                    float dagtotaal = uren.geefDagtotaal();
                    if (project.equals("verlof") || (!project.equals("verlof") && totaal > 0)) {
                        granttotal += totaal;
                        String tekst = String.format("file: %s, minuten: %6.1f, uren: %3.1f, dagen: %4.2f, uren gewerkt: %2.0f \n", xlsFile, (float) totaal, (float) totaal / 60, (float) totaal / 60 / uren.URENPERDAG, dagtotaal);
                        listModel.addElement(tekst);
                    }
                    uren.sluitWerkblad();
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

    public static void main(String[] args) {

        // initialiseer logger
        Logger log = Logger.getLogger(MijnUren.class.getName());

        String logDir = ".";
        String logNaam = "MijnUren.log";
        try {
            MijnLog mijnlog = new MijnLog(logDir, logNaam, true);
            log = mijnlog.getLog();
            log.setLevel(Level.INFO);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

        JFrame frame = new JFrame("MijnUren");
        frame.setContentPane(new MijnUren().mainPanel);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        frame.setSize();
        frame.pack();
        frame.setVisible(true);

    }



}
