import annotation.ImportExcel;

public class Bean {
    @ImportExcel(value = "用户名")
    private String username;
    @ImportExcel(value = "年龄")
    private double age;

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public double getAge() {
        return age;
    }

    public void setAge(double age) {
        this.age = age;
    }
}
