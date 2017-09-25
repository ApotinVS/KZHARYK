
import java.io.IOException;
import java.nio.file.*;
import java.util.ArrayList;

public class DirectoryFileNames {
    static ArrayList<String> fileNames = new ArrayList<String>();
    public static ArrayList<String> GetFileNames(String path){
        try (DirectoryStream<Path> stream = Files.newDirectoryStream(Paths.get(path))) {
            for (Path file: stream) {
                if(!file.toFile().isDirectory() ) {
                    fileNames.add(file.getFileName().toString());

                }
            }
        } catch (IOException | DirectoryIteratorException x) {
            System.err.println(x);
        }
        return fileNames;

    }
}