import org.junit.Test;

/**
 * Created by chenlian on 14-8-30.
 */
public class ToFile {
    @Test
    public void toExcel(){
        MicroSoftDocument document=new Excel();
        if (document.toFile("E:\\AMD\\book.xls")){
         System.out.println("导出成功");
        }
    }

    @Test
    public void tts(){

        MicroSoftDocument document=new Excel();
        document.openFile("E:\\AMD\\book.xls");


    }
}
