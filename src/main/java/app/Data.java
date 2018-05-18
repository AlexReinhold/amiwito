package app;

public class Data {

    public static String WHATSAPP = "Whatsapp";
    public static String NO_WHATSAPP = "NO Whatsapp";

    private int id;
    private String number;
    private String type;

    public Data(int id, String number) {
        this.id = id;
        this.number = number;
    }

    public int getId() {
        return id;
    }

    public Data setId(int id) {
        this.id = id;
        return this;
    }

    public String getNumber() {
        return number;
    }

    public Data setNumber(String number) {
        this.number = number;
        return this;
    }

    public String getType() {
        return type;
    }

    public Data setType(String type) {
        this.type = type;
        return this;
    }

}
