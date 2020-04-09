import org.h2.tools.DeleteDbFiles;

import java.io.IOException;
import java.sql.*;

public class Main {
    public static void main(String[] args) throws SQLException, IOException {
        DeleteDbFiles.execute("~", "test", true);

        try {
            Class.forName("org.h2.Driver");
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        }


            Connection conn = DriverManager.getConnection("jdbc:h2:~/test");
            Statement stat = conn.createStatement();

            stat.execute("CREATE TABLE test(id bigint primary key, name varchar(255), surName varchar(255), birthDay date )");
            stat.execute("INSERT INTO test VALUES(1, 'Rodion', 'Yesipov', '1993-08-21')");
            stat.execute("INSERT INTO test VALUES(2, 'Vasya', 'Pupkin', '1990-08-01')");
            ResultSet rs;
            rs = stat.executeQuery("SELECT * FROM test");
            ExcelWrite excelWrite = new ExcelWrite();
            excelWrite.resultSetToFile(rs, "D:\\1.xls");


    }
}
