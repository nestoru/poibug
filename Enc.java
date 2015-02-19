import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.OutputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;

public class Enc {
  public static void main(String[] args) throws Exception{

    if(args.length != 2) {
      System.out.println("Usage: java -cp \"*\" Enc <inputFilePath> <outputFilePath>");
      return;
    }

    String inputFilePath = args[0];
    String outputFilePath = args[1];

    POIFSFileSystem fs = new POIFSFileSystem();
    EncryptionInfo info = new EncryptionInfo(fs, EncryptionMode.agile);
    // EncryptionInfo info = new EncryptionInfo(fs, EncryptionMode.agile, CipherAlgorithm.aes192, HashAlgorithm.sha384, -1, -1, null);

    Encryptor enc = info.getEncryptor();
    enc.confirmPassword("password");

    OPCPackage opc = OPCPackage.open(inputFilePath, PackageAccess.READ_WRITE);
    OutputStream os = enc.getDataStream(fs);
    opc.save(os);
    opc.close();
    FileOutputStream fos = new FileOutputStream(outputFilePath);
    fs.writeFilesystem(fos);
    fos.close();

    //XSSFWorkbook wb = new XSSFWorkbook(inputFilePath);
    //OutputStream os = enc.getDataStream(fs);
    //File tempFile = File.createTempFile("poi-enc", ".tmp");
    //System.out.println("Using temporary file '" + tempFile.getAbsolutePath() + "'");
    //wb.write(new FileOutputStream(tempFile));

  }
}
