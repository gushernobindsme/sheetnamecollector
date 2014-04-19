import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * エクセルファイルのシート名集計ツール。
 * <pre>
 * 指定したフォルダからエクセルファイルを再帰的に検索し、
 * シート名を集計してテキストファイルに出力します。
 * 
 * 設定する引数は以下の通り。
 * ・第一引数：検索対象のディレクトリパス
 * ・第二引数：検索結果ファイルの保存先パス
 * </pre>
 * 
 * @author TakumiEra
 * 
 */
public class Main {

	/**
	 * ディレクトリ名のリスト。
	 */
	static List<String> DIR_LIST = new ArrayList<String>();
	
	/**
	 * シート名のリスト。
	 */
	static List<String> SHEET_NAME_LIST = new ArrayList<String>();

	/**
	 * 本処理。
	 * 
	 * @param args
	 * @throws Exception
	 */
	public static void main(String[] args) throws Exception {
		if (args.length == 0) {
			System.out.println("引数を指定してください。");
			return;
		}
		if(StringUtils.isEmpty(args[0])){
			System.out.println("引数が不足しています。");
			return;
		}
		if(args.length == 1 || StringUtils.isEmpty(args[1])){
			System.out.println("引数が不足しています。");
			return;
		}
		System.out.println("実行中...");

		// 引数から情報を取得
		String searchDirPath = args[0];
		String saveDirPath = args[1];

		// 検索の実行
		searchDir(searchDirPath);

		// 検索結果をテキストファイルに書き込み
		write(saveDirPath);
	}

	/**
	 * 指定したパスを検索する。
	 * 
	 * @param aPath ディレクトリパス
	 * @return
	 * @throws Exception
	 */
	private static boolean searchDir(final String aPath) throws Exception {
		File file = new File(aPath);
		File[] listFiles = file.listFiles(new FilenameFilter() {

			public boolean accept(File aDir, String aName) {
				// ドットで始まるファイルは対象外
				if (aName.startsWith(".")) {
					return false;
				}

				// 対象要素の絶対パスを取得
				String absolutePath = aDir.getAbsolutePath() + File.separator + aName;

				// エクセルファイルのみ対象とする。
				if (new File(absolutePath).isFile()
						&& (absolutePath.endsWith(".xls") || absolutePath
								.endsWith(".xlsx"))) {
					return true;
				} else {
					// ディレクトリの場合、再び同一メソッドを呼出す。
					try {
						return searchDir(absolutePath);
					} catch (Exception e) {
						e.printStackTrace();
						return false;
					}
				}
			}
		});

		if (listFiles == null) {
			return false;
		}

		// シート名を取得
		for (File f : listFiles) {
			if (f.isFile()) {
				addSheetName(f);
			}
		}
		return true;
	}

	/**
	 * 対象のエクセルシートからシート名を取得し、リストに格納します。
	 * 
	 * @param aFile ファイルオブジェクト
	 * @throws FileNotFoundException
	 */
	private static void addSheetName(File aFile) throws FileNotFoundException {
		// Excelファイルの読込み
		InputStream inputStream = new FileInputStream(aFile);
		Workbook workBook;

		try {
			workBook = WorkbookFactory.create(inputStream);
		} catch (Exception ex) {
			return;
		}

		// シート枚数を読込み
		int numberOfSheets = workBook.getNumberOfSheets();

		// シート枚数分ループ処理
		for (int i = 0; i < numberOfSheets; i++) {
			// シート情報を取得
			Sheet sheet = workBook.getSheetAt(i);
			// シート名を取得
			String sheetName = sheet.getSheetName();

			// 取得した情報を書き込み
			DIR_LIST.add(aFile.getAbsolutePath());
			SHEET_NAME_LIST.add(sheetName);
			
			System.out.println("Path: " + aFile.getAbsolutePath() + " SheetName: " + sheetName);
		}

	}
	
	/**
	 * テキストファイルへの書き込みを行います。
	 * 
	 * @param path ディレクトリパス
	 */
	private static void write(final String path) {
		try {
			File file = new File(path + "\\" + "result.csv");

			if (checkBeforeWritefile(file)) {
				BufferedWriter bufferedWriter = new BufferedWriter(new FileWriter(file));

				for (int i = 0; i < DIR_LIST.size(); i++) {
					bufferedWriter.write(DIR_LIST.get(i));
					bufferedWriter.write(",");
					bufferedWriter.write(SHEET_NAME_LIST.get(i));
					bufferedWriter.newLine();
				}

				bufferedWriter.close();
			} else {
				System.out.println("ファイルに書き込めません");
			}
		} catch (IOException e) {
			System.out.println(e);
		}
	}

	/**
	 * ファイルへの書き込みが可能かどうかをチェックします。
	 * 
	 * @param aFile ファイルオブジェクト
	 * @return
	 */
	private static boolean checkBeforeWritefile(File aFile) {
		if (aFile.exists()) {
			if (aFile.isFile() && aFile.canWrite()) {
				System.out.println(aFile.getAbsolutePath() + "に上書きします。");
				return true;
			}else{
				return false;
			}
		}
		System.out.println(aFile.getAbsolutePath() + "を作成します。");
		return true;
	}
}
