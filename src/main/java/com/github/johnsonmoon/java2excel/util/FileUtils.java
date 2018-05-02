package com.github.johnsonmoon.java2excel.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Created by xuyh at 2018/1/2 13:37.
 */
public class FileUtils {
	/**
	 * Judge whether the given file by filePathName exists.
	 *
	 * @param filePathName file path name
	 * @return true/false
	 */
	public static boolean exists(String filePathName) {
		File file = new File(filePathName);
		return file.exists();
	}

	/**
	 * Delete file, if not exist return true;
	 *
	 * @param filePathName file path name
	 * @return true/false
	 */
	public static boolean delete(String filePathName) {
		File file = new File(filePathName);
		if (file.exists()) {
			try {
				return file.delete();
			} catch (Exception e) {
				return false;
			}
		} else {
			return true;
		}
	}

	/**
	 * Get file name from file path name.
	 *
	 * @param filePathName file path name
	 * @return file name
	 */
	public static String getName(String filePathName) {
		if (filePathName == null || filePathName.isEmpty())
			return "";
		if (!filePathName.contains(File.separator))
			return filePathName;
		return filePathName.substring(filePathName.lastIndexOf(File.separator) + 1, filePathName.length());
	}

	/**
	 * Get file path from file path name.
	 *
	 * @param filePathName file path name
	 * @return file name
	 */
	public static String getPath(String filePathName) {
		if (filePathName == null || filePathName.isEmpty())
			return "";
		if (!filePathName.contains(File.separator))
			return "";
		return filePathName.substring(0, filePathName.lastIndexOf(File.separator));
	}

	/**
	 * Add head string into file name.
	 * <p>
	 * <pre>
	 *     D:\\test.txt   temp_
	 *     --->
	 *     D:\\temp_test.txt
	 * </pre>
	 *
	 * @param filePathName file path name
	 * @param head         tail string
	 * @return modified string
	 */
	public static String addHead2Name(String filePathName, String head) {
		if (filePathName == null || filePathName.isEmpty())
			return head + "";
		String formal = getPath(filePathName);
		String behind = getName(filePathName);
		return formal + File.separator + head + behind;
	}

	/**
	 * Add tail string into file name.
	 * <p>
	 * <pre>
	 *     D:\\test.txt   _temp
	 *     --->
	 *     D:\\test_temp.txt
	 * </pre>
	 *
	 * @param filePathName file path name
	 * @param tail         tail string
	 * @return modified string
	 */
	public static String addTail2Name(String filePathName, String tail) {
		if (filePathName == null || filePathName.isEmpty())
			return "" + tail;
		if (!filePathName.contains("."))
			return filePathName + tail;
		String formal = filePathName.substring(0, filePathName.lastIndexOf("."));
		String behind = filePathName.substring(filePathName.lastIndexOf("."), filePathName.length());
		return formal + tail + behind;
	}

	/**
	 * Copy file {filePathName} into a new file {newFilePathName}. Delete and create new if {newFilePathName} exists.
	 *
	 * @param filePathName    to be copy
	 * @param newFilePathName target file
	 * @return true/false
	 */
	public static boolean copyFile(String filePathName, String newFilePathName) {
		boolean flag = true;
		File file = new File(filePathName);
		if (!file.exists())
			return false;
		FileOutputStream fileOutputStream = null;
		InputStream inputStream = null;
		try {
			File copy = createFileDE(newFilePathName);
			if (copy == null)
				return false;
			fileOutputStream = new FileOutputStream(copy);
			inputStream = new FileInputStream(file);
			byte[] b = new byte[1024];
			int readLength;
			while ((readLength = inputStream.read(b)) != -1) {
				fileOutputStream.write(b, 0, readLength);
			}
		} catch (Exception e) {
			flag = false;
		} finally {
			if (fileOutputStream != null) {
				try {
					fileOutputStream.close();
				} catch (Exception e) {
					flag = false;
				}
			}
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (Exception e) {
					flag = false;
				}
			}
		}
		return flag;
	}

	/**
	 * Rename file.
	 *
	 * @param oldFilePathName old file path name
	 * @param newFilePathName new file path name
	 * @return true/false
	 */
	public static boolean rename(String oldFilePathName, String newFilePathName) {
		File file = new File(oldFilePathName);
		if (!file.exists())
			return false;
		return file.renameTo(new File(newFilePathName));
	}

	/**
	 * Create new file with its parent directory, delete and create new one if exists.
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
	 * Open file with its parent directory, create a new one if it isn't exist.
	 *
	 * @param filePathName file path name
	 * @return null if open file failed
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
