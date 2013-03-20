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

    public int getTijd_in() {
        return tijd_in;
    }

    public int getTijd_uit() {
        return tijd_uit;
    }
}
