package mutilExcel;

public class ReportParameterX {
    private String propertyName;

    private String propertyValue;

    public ReportParameterX(String propertyName, String propertyValue) {
        this.propertyName = propertyName;
        this.propertyValue = propertyValue;
    }

    public String getPropertyName() {
        return this.propertyName;
    }

    public String getPropertyValue() {
        return this.propertyValue;
    }

    public void setPropertyName(String propertyName) {
        this.propertyName = propertyName;
    }

    public void setPropertyValue(String propertyValue) {
        this.propertyValue = propertyValue;
    }
}
