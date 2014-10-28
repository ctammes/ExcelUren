package nl.ctammes.exceluren;

/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 16:40
 * To change this template use File | Settings | File Templates.
 */
public class Verlofdag {

    private int dag;
    private String datum;
    private float minuten;

    public Verlofdag(int dag, String datum, float minuten) {
        this.dag = dag;
        this.datum = datum;
        this.minuten = minuten;
    }

    public int getDag() {
        return dag;
    }

    public String getDagnaam() {
        String dagnaam = "";
        switch (dag) {
            case 2:
                dagnaam = "maandag";
                break;
            case 3:
                dagnaam = "dinsdag";
                break;
            case 4:
                dagnaam = "woensdag";
                break;
            case 5:
                dagnaam = "donderdag";
                break;
            case 6:
                dagnaam = "vrijdag";
                break;
            case 7:
                dagnaam = "zaterdag";
                break;
            case 8:
                dagnaam = "zondag";
                break;
        }
        return dagnaam;
    }

    public String getDagnaamKort() {
        return getDagnaam().substring(0,2);
    }

    public String getDatum() {
        return datum;
    }

    public float getMinuten() {
        return minuten;
    }
}
