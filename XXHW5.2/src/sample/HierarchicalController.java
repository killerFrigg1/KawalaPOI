package sample;

/**
 * Created by pwilkin on 30-Nov-17.
 */
public interface HierarchicalController<P extends HierarchicalController<?>> {

    public P getParentController();
    public void setParentController(P parent);

}
