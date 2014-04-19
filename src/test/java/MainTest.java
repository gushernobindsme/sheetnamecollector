import org.junit.Test;

public class MainTest {
	
	@SuppressWarnings("static-access")
	@Test
	public void main正常系テスト() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp","C:\\temp"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main異常系テスト1() throws Exception {
		// パラメータ作成
		String[] args = {};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main異常系テスト2() throws Exception {
		// パラメータ作成
		String[] args = {"C:\\temp"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
	@SuppressWarnings("static-access")
	@Test
	public void main異常系テスト3() throws Exception {
		// パラメータ作成
		String[] args = {"","C:\\temp"};
		
		// テスト実行
		Main target = new Main();
		target.main(args);
	}
	
}
