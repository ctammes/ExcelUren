/**
 * Created with IntelliJ IDEA.
 * User: chris
 * Date: 20-3-13
 * Time: 16:40
 * To change this template use File | Settings | File Templates.
 */
public class Werkdag {

    private int dag;
    private int tijd_in;
    private int tijd_uit;

    public Werkdag(int dag, int tijd_in, int tijd_uit) {
        this.dag = dag;
        this.tijd_in = tijd_in;
        this.tijd_uit = tijd_uit;
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

    public int getTijd_in() {
        return tijd_in;
    }

    public int getTijd_uit() {
        return tijd_uit;
    }
}
