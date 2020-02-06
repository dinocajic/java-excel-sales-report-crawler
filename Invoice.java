import java.util.ArrayList;
import java.util.List;

public class Invoice {

    public static int number_of_items = 0;
    public static double static_total = 0;
    public static double static_cost = 0;
    public static double static_additional_charges = 0;

    private String invoice_number, type, account, name, job_num, sales_user, sis_mat_br, date, total, cost, gm_perc;

    private List<String[]> items = new ArrayList<>();

    public String getInvoice_number() { return invoice_number; }

    public String getType() { return type; }

    public String getAccount() { return account; }

    public String getName() { return name; }

    public String getJob_num() { return job_num; }

    public String getSales_user() { return sales_user; }

    public String getSis_mat_br() { return sis_mat_br; }

    public String getDate() { return date; }

    public String getTotal() { return total; }

    public String getCost() { return cost; }

    public String getGm_perc() { return gm_perc; }

    public List<String[]> getItems() { return items; }

    public void setInvoice_number(String invoice_number) {  this.invoice_number = invoice_number; }

    public void setType(String type) { this.type = type; }

    public void setAccount(String account) { this.account = account; }

    public void setName(String name) { this.name = name; }

    public void setJob_num(String job_num) { this.job_num = job_num; }

    public void setSales_user(String sales_user) { this.sales_user = sales_user; }

    public void setSis_mat_br(String sis_mat_br) { this.sis_mat_br = sis_mat_br; }

    public void setDate(String date) { this.date = date; }

    public void setTotal(String total) { this.total = total; }

    public void setCost(String cost) { this.cost = cost; }

    public void setGm_perc(String gm_perc) { this.gm_perc = gm_perc; }

    public void setItems(List<String[]> items) {
        this.items = items;
        this.update( items );
    }

    public void update( List<String[]> items ) {

        double total = 0;
        double cost  = 0;

        for ( String[] item : items ) {
            double item_qty   = Double.parseDouble( item[2] );
            double item_total = Double.parseDouble( item[4] );
            double item_cost  = Double.parseDouble( item[6] );

            total += item_qty * item_total;
            cost  += item_qty * item_cost;
        }

        double additional_charges = Double.parseDouble( this.getTotal() ) - total;

        if ( additional_charges > 0.25 ) {
            this.items.add( new String[]{
                    "Extra",
                    "Additional Charges such as FRT or CC",
                    "",
                    "",
                    String.format("%.2f", additional_charges),
                    "",
                    "¯\\_(ツ)_/¯",
                    ""
            } );

            static_additional_charges += additional_charges;
        }

        double difference = total - cost;
        double gm_percent = difference/total * 100;

        this.setGm_perc( String.format("%.2f", gm_percent) );
        this.setCost( String.format("%.2f", cost) );
        this.setTotal( String.format("%.2f", total) );

        static_total += total;
        static_cost  += cost;
    }

    @Override
    public String toString() {

        String part_nums = "";

        for ( String[] item : items ) {
            part_nums = part_nums + item[0] + " :: ";
        }

        return "Invoice{" +
                "invoice_number='" + invoice_number + '\'' +
                ", type='" + type + '\'' +
                ", account='" + account + '\'' +
                ", name='" + name + '\'' +
                ", job_num='" + job_num + '\'' +
                ", sales_user='" + sales_user + '\'' +
                ", sis_mat_br='" + sis_mat_br + '\'' +
                ", date='" + date + '\'' +
                ", total='" + total + '\'' +
                ", cost='" + cost + '\'' +
                ", gm_perc='" + gm_perc + '\'' +
                ", items=" + part_nums +
                '}';
    }
}
