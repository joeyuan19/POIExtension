/* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * 
 * Copyright 2013 Joseph Yuan
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 * 
 * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * */



package excelUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;

import org.apache.tika.Tika;


public class FileUtils {
	public static String getPWD() {
		return System.getProperty("user.dir") + File.separator;
	}
	/* Excel extension tools */
	public static String getExt(File file) {
		return getExt(file.getName());
	}
	public static String getExt(String filename) {
		int i;
		String buf = "";
		boolean noExtFound = true;
		for (i = 0; i < filename.length(); i++) {
			if (filename.charAt(i) == '.') {
				buf = "";
				noExtFound = false;
			}
			buf += filename.charAt(i);
		}
		if (noExtFound) {
			return null;
		}
		return buf;
	}
	public static String parseExt(File file) {
		return parseExt(file.getName());
	}

	public static String parseExt(String filename) {
		int i;
		String buf = "", parsedFilename = "";
		boolean noExtFound = true;
		for (i = 0; i < filename.length(); i++) {
			if (filename.charAt(i) == '.' && i > 0) {
				parsedFilename += buf;
				buf ="";
				noExtFound = false;
			}
			buf += filename.charAt(i);
		}
		if (noExtFound) {
			parsedFilename += buf;
		}
		return parsedFilename;
	}
	public static boolean hasProperExt(String filename) {
		String ext = getExt(filename);
		if (ext == null) {
			return false;
		}
		if (ext.equals(".xls") || ext.equals(".xlsx") ) {
			return true;
		}
		return false;
	}
	/* File Location Methods */
	public static File locateAndOpenFile(String filename) throws Exception {
		return locateAndOpenFile(filename,getPWD(),0,false,true,false,null);
	}
	public static File locateAndOpenFile(String filename,String path,String msg) throws Exception {
		return locateAndOpenFile(filename,path,0,false,true,false,msg);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,String msg) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,true,false,msg);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,boolean recursive,String msg) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,recursive,false,msg);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,boolean recursive,boolean makedirs,String msg) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,recursive,makedirs,msg);
	}
	public static File locateAndOpenFile(String filename, String path, int level,boolean matchPartial,boolean recursive,boolean makedirs, String msg) throws Exception {
		File locatedFile = null;
		String curFileName;
		File dir = new File(path);
		if (!dir.exists()) {return null;}
		ArrayList<File> subDirs = new ArrayList<File>(); 
		dir.mkdirs();
		if (level == 0) {
			filename = parseExt(filename);
		}
		for (File curFile: dir.listFiles()) {
			if (curFile.exists()) {
				if (recursive && curFile.isDirectory()) {
					subDirs.add(curFile);
				}
				if (curFile.isFile()) {
					curFileName = parseExt(curFile.getName());
					if (matchPartial) {
						if (curFileName.toLowerCase().contains(filename.toLowerCase()) || 
								filename.toLowerCase().contains(curFileName.toLowerCase())) {
							locatedFile = curFile;
							break;
						}
					} else {
						if (curFileName.equals(filename)) {
							locatedFile = curFile;
							break;
						}
					}
				}
			}
		}
		if (locatedFile == null && level == 0) {
			/* Manual Location required */
			locatedFile = manualLocate("Locate " + filename + "...");
		} else if (locatedFile == null) {
			for (File curFile: subDirs) { 
				locatedFile = locateAndOpenFile(filename,curFile.getAbsolutePath(),level+1,matchPartial,recursive,makedirs,msg);
				if (locatedFile != null) {
					break;
				}
			}
		}
		return locatedFile;
	}
	/* File location by REGEX */
	public static File locateAndOpenFileRegex(String regex) {
		return locateAndOpenFileRegex(regex,getPWD(),true,"Locate file matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex,boolean recursive) {
		return locateAndOpenFileRegex(regex,getPWD(),recursive,"Locate file matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex,boolean recursive, String msg) {
		return locateAndOpenFileRegex(regex,getPWD(),recursive,msg,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory) {
		return locateAndOpenFileRegex(regex,startingDirectory,true,"Locate file matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, String msg) {
		return locateAndOpenFileRegex(regex,startingDirectory,true,msg,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, boolean recursive) {
		return locateAndOpenFileRegex(regex,startingDirectory,recursive,"Locate file matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, boolean recursive, String msg, boolean offerOptOut) {
		Pattern pattern = Pattern.compile(regex,Pattern.CASE_INSENSITIVE);
		return locateAndOpenFileRegex(pattern,startingDirectory,recursive,msg,0,offerOptOut);
	}
	public static File locateAndOpenFileRegex(Pattern pattern, String startingDirectory, boolean recursive, String msg,int level,boolean offerOptOut) {
		File locatedFile = null;
		String curFileName;
		File dir = new File(startingDirectory);
		if (!dir.exists()) {return null;}
		ArrayList<File> subDirs = new ArrayList<File>(); 
		dir.mkdirs();
		Matcher m;
		for (File curFile: dir.listFiles()) {
			if (curFile.exists()) {
				if (curFile.isFile()) {
					curFileName = curFile.getName();
					m = pattern.matcher(curFileName);
					if (m.find()) {
						locatedFile = curFile;
						break;
					}
				}
				if (recursive && curFile.isDirectory()) {
					subDirs.add(curFile);
				}
			}
		}
		if (locatedFile == null) {
			for (File curFile: subDirs) {
				locatedFile = locateAndOpenFileRegex(pattern,curFile.getAbsolutePath(),recursive,msg,level+1,offerOptOut);
				if (locatedFile != null) {
					break;
				}
			}
		}
		if (locatedFile == null && level == 0) {
			int manLocate = JOptionPane.YES_OPTION;
			if (offerOptOut) {
				Object[] objs = {"Locate","Skip"};
				manLocate = JOptionPane.showOptionDialog(null, "This report was not found, would you like to locate the report manually?", "File Not Found", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE, null, objs, objs[0]);
			}
			if (manLocate == JOptionPane.YES_OPTION) {
				locatedFile = manualLocate(startingDirectory,msg);
			}
		}
		return locatedFile;
	}
	/* Manual Location */
	public static File manualLocate(String msg) {
		return manualLocate(msg, false);
	}
	public static File manualLocate(String msg, boolean isDir) {
		return manualLocate(System.getProperty("user.dir"),msg,isDir);
	}
	public static File manualLocate(String dir, String msg) {
		return manualLocate(dir,msg,false);
	}
	public static File manualLocate(String dir, String msg, boolean isDir) {
		return manualLocate(dir,msg,isDir,(JFrame)null);
	}
	public static File manualLocate(String dir, String msg, boolean isDir, JFrame frame) {
		File located = null;
		if (!(new File(dir)).exists()) dir = getPWD();
		JFileChooser locator = new JFileChooser(dir);
		if (isDir) {
			locator.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		}
		locator.setDialogTitle(msg);
		int returnVal = locator.showOpenDialog(frame);
		if (returnVal == JFileChooser.APPROVE_OPTION) {
			located = locator.getSelectedFile();
		} else {
			located = null;
		}
		return located;
	}
	/* Folder search */
	public static File locateFolder(String folderName) {
		return locateFolder(folderName,getPWD(),true,false,true);
	}
	public static File locateFolder(String folderName, String startingDir) {
		return locateFolder(folderName,startingDir,true,false,true);
	}
	public static File locateFolder(String folderName, File startingDir) {
		return locateFolder(folderName,startingDir.getName(),true,false,true);
	}
	public static File locateFolder(String folderName, String startingDir, boolean recursive, boolean matchPartial, boolean ignoreCase) {
		File itr = new File(startingDir);
		if (!itr.exists()) {
			return null;
		}
		return locateFolder(folderName,itr,recursive,matchPartial,ignoreCase);
	}
	public static File locateFolder(String folderName, File startingDir, boolean recursive, boolean matchPartial, boolean ignoreCase) {
		ArrayList<File> dirsToSearch = new ArrayList<File>();
		for (File curDir : startingDir.listFiles()) {
			if (curDir.isDirectory()) {
				if (ignoreCase ?
						matchPartial ? curDir.getName().toLowerCase().contains(folderName.toLowerCase()) : curDir.getName().equalsIgnoreCase(folderName)
								:
									matchPartial ? curDir.getName().contains(folderName) : curDir.getName().equals(folderName)) {
					return curDir;
				}
				if (recursive) {
					dirsToSearch.add(curDir);
				}
			}
		}
		File folder;
		for (File itr : dirsToSearch) { 
			if ((folder = locateFolder(folderName,itr,recursive,matchPartial,ignoreCase)) != null) {
				return folder;
			}
		}
		return null;
	}

	/* File Type detection */
	public static String detectFileType(File file) throws IOException {
		return detectFileType(file.getAbsolutePath());
	}
	public static String detectFileType(String path) throws IOException {
		Tika tika = new Tika();
		String type = tika.detect(path);
		return type;

	}
	/* Ensured File deletion */
	public static boolean deleteFile(File file) {
		boolean status = true;
		if (file != null && file.exists()) {
			status = file.delete();
			if (!status) {
				try {
					org.apache.commons.io.FileUtils.forceDelete(file);
					status = true;
				} catch (Exception e) {
					e.printStackTrace();
					status = false;
				}
				if (!status) {
					status = org.apache.commons.io.FileUtils.deleteQuietly(file);
				}	
			}
		}
		return status;
	}

	/* File path manipulation */
	public static String joinPath(String...paths) {
		return joinPath(false,false,paths);
	}
	public static String joinPath(boolean relativePath, String...paths ){
		return joinPath(false,relativePath,paths);
	}
	public static String joinPath(boolean trailingSeparator, boolean relativePath, String... paths) {
		String path = "";
		for (int i = 0; i < paths.length; i++) {
			if (paths[i].length() > 0) {
				path += (paths[i].startsWith(File.separator) ? (i == 0 && !relativePath ? paths[i] : paths[i].substring(1)) : paths[i]) +
						(trailingSeparator ? (paths[i].endsWith(File.separator) ? "" : File.separator) : (i != paths.length-1 && !paths[i].endsWith(File.separator)) ? File.separator : ""); // Optional trailing separator
			}
		}
		return path;
	}
	public static String shortenPath(String path) {
		return shortenPath(path,1);
	}
	public static String shortenPath(String path, int distance) {
		int i = 0, L = path.length(), index = L, count = 0;
		for (i = L-1; i >= 0; i--) {
			if (path.charAt(i) == File.separatorChar) {
				count++;
				index = i+1;
			}
			if (count >= distance) {
				break;
			}
		}
		return path.substring(index);
	}
	public static FileOutputStream getFileOutputStream(String path) throws IOException {
		return getFileOutputStream(path,false,false);
	}
	public static FileOutputStream getFileOutputStream(String path, boolean create) throws IOException {
		return getFileOutputStream(path,create,false);
	}
	public static FileOutputStream getFileOutputStream(String path, boolean create, boolean append) throws IOException {
		File file = new File(path);
		if (!file.exists() && create) {
			file.createNewFile();
		} else if (!file.exists()) {
			return null;
		}
		FileOutputStream out = new FileOutputStream(path,append);
		return out;
	}
	public static FileInputStream getFileInputStream(String path) throws IOException {
		return getFileInputStream(path,false);
	}
	public static FileInputStream getFileInputStream(String path, boolean create) throws IOException {
		File file = new File(path);
		if (!file.exists() && create) {
			file.createNewFile();
		} else if (!file.exists()) {
			return null;
		}
		FileInputStream in = new FileInputStream(path);
		return in;
	}

	
	
	/* Main Method for tests */
	public static void main(String args[]) {
		System.out.println();
	}
}
