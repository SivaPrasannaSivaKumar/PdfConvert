package convertion;

import org.apache.commons.io.FilenameUtils;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.PageOrientationType;
import com.aspose.cells.Workbook;

public class Pdf {

	public static void main(String[] args) throws Exception {
		String file = "D:\\TE411917.xls";
		Workbook workbook = new Workbook(file);
		String path = FilenameUtils.getPath(file);
		String fileName = FilenameUtils.getBaseName(file);
		for(int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			workbook.getWorksheets().get(i).getPageSetup().setOrientation(PageOrientationType.LANDSCAPE);
		}
		System.out.println(fileName);
		workbook.save(path + fileName + ".pdf", FileFormatType.PDF);
	}

}
