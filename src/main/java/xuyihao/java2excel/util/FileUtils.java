package xuyihao.java2excel.util;

import java.io.File;

/**
 * Created by xuyh at 2018/1/2 13:37.
 */
public class FileUtils {
	/**
	 * create new file with its parent directory, delete and create new one if exists
	 *
	 * @param filePathName file path name
	 * @return null if create file failed
	 * @throws Exception
	 */
	public static File createFileDE(String filePathName) throws Exception {
		File file = new File(filePathName);
		if (file.exists()) {
			if (!file.delete())
				return null;
		}
		if (!makeDir(file.getParentFile()))
			return null;
		if (!file.createNewFile())
			return null;
		return file;
	}

	/**
	 * create new file with its parent directory
	 *
	 * @param filePathName file path name
	 * @return null if create file failed
	 * @throws Exception
	 */
	public static File createFile(String filePathName) throws Exception {
		File file = new File(filePathName);
		if (!file.exists()) {
			if (!makeDir(file.getParentFile()))
				return null;
			if (!file.createNewFile())
				return null;
		}
		return file;
	}

	/**
	 * create directory recursively
	 *
	 * @param dir
	 * @return
	 * @throws Exception
	 */
	public static boolean makeDir(File dir) throws Exception {
		if (!dir.exists()) {
			File parent = dir.getParentFile();
			if (parent != null)
				makeDir(parent);
			return dir.mkdir();
		}
		return true;
	}
}
