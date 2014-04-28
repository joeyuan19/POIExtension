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
		int n = filename.lastIndexOf('.');
		if (n > 0) {
			return filename.substring(n+1);
		}
		return null;
	}
	public static String parseExt(File file) {
		return parseExt(file.getName());
	}
	public static String parseExt(String filename) {
		int n = filename.lastIndexOf('.');
		if (n > 0) {
			return filename.substring(n+1);
		}
		return filename;
	}
	public static String removeExt(File file) {
		return parseExt(file);
	}
	public static String removeExt(String filename) {
		return parseExt(filename);
	}
	public static boolean hasProperExt(String filename) {
		return filename.toLowerCase().endsWith(".xls") || filename.toLowerCase().endsWith(".xlsx");
	}
	/* File Location Methods */
	public static File locateAndOpenFile(String filename) throws Exception {
		return locateAndOpenFile(filename,getPWD(),0,false,true,false,null,true);
	}
	public static File locateAndOpenFile(String filename,String path,String slug) throws Exception {
		return locateAndOpenFile(filename,path,0,false,true,false,slug,true);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,String slug) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,true,false,slug,true);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,String slug,boolean offerOptOut) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,true,false,slug,offerOptOut);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,boolean recursive,String slug) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,recursive,false,slug,true);
	}
	public static File locateAndOpenFile(String filename,String path,boolean matchPartial,boolean recursive,boolean makedirs,String slug) throws Exception {
		return locateAndOpenFile(filename,path,0,matchPartial,recursive,makedirs,slug,true);
	}
	public static File locateAndOpenFile(String filename, String path, int level,boolean matchPartial,boolean recursive,boolean makedirs, String slug,boolean offerOptOut) throws Exception {
		String msg = "Locate correct '" + slug + "' file";
		File locatedFile = null;
		String curFileName;
		File dir = new File(path);
		if (dir.exists()) {
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
			if (locatedFile == null) {
				for (File curFile: subDirs) { 
					locatedFile = locateAndOpenFile(filename,curFile.getAbsolutePath(),level+1,matchPartial,recursive,makedirs,slug,offerOptOut);
					if (locatedFile != null) {
						break;
					}
				}
			}
		}
		if (locatedFile == null && level == 0) {
			/* Manual Location required */
			int manLocate = JOptionPane.YES_OPTION;
			if (offerOptOut) {
				Object[] objs = {"Locate","Skip"};
				manLocate = JOptionPane.showOptionDialog(null, "Could not automatically find " + slug + "file\nwould you like to locate the report manually?",
						"File Not Found", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE, null, objs, objs[0]);
			}
			if (manLocate == JOptionPane.YES_OPTION) {
				locatedFile = manualLocate(path,msg);
			}
		}
		return locatedFile;
	}
	/* File location by REGEX */
	public static File locateAndOpenFileRegex(String regex) {
		return locateAndOpenFileRegex(regex,getPWD(),true,"File matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex,boolean recursive) {
		return locateAndOpenFileRegex(regex,getPWD(),recursive,"File matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex,boolean recursive, String slug) {
		return locateAndOpenFileRegex(regex,getPWD(),recursive,slug,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory) {
		return locateAndOpenFileRegex(regex,startingDirectory,true,"File matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, String slug) {
		return locateAndOpenFileRegex(regex,startingDirectory,true,slug,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, boolean recursive) {
		return locateAndOpenFileRegex(regex,startingDirectory,recursive,"File matching pattern: " + regex,false);
	}
	public static File locateAndOpenFileRegex(String regex, String startingDirectory, boolean recursive, String slug, boolean offerOptOut) {
		Pattern pattern = Pattern.compile(regex,Pattern.CASE_INSENSITIVE);
		return locateAndOpenFileRegex(pattern,startingDirectory,recursive,slug,0,offerOptOut);
	}
	public static File locateAndOpenFileRegex(Pattern pattern, String startingDirectory, boolean recursive, String slug,int level,boolean offerOptOut) {
		String msg = "Could not locate correct '" + slug + "' file.";
		File locatedFile = null;
		String curFileName;
		File dir = new File(startingDirectory);
		if (dir.exists()) {
			ArrayList<File> subDirs = new ArrayList<File>(); 
			dir.mkdirs();
			Matcher m;
			for (File curFile: dir.listFiles()) {
				if (curFile.exists()) {
					if (curFile.isFile()) {
						curFileName = curFile.getName();
						m = pattern.matcher(curFileName);
						if (m.matches()) {
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
					locatedFile = locateAndOpenFileRegex(pattern,curFile.getAbsolutePath(),recursive,slug,level+1,offerOptOut);
					if (locatedFile != null) {
						break;
					}
				}
			}
		}
		if (locatedFile == null && level == 0) {
			int manLocate = JOptionPane.YES_OPTION;
			if (offerOptOut) {
				Object[] objs = {"Locate","Skip"};
				manLocate = JOptionPane.showOptionDialog(null, "Unable to automatically find correct " + slug 
						+ " file\nWould you like to locate the report manually?", 
						"File Not Found", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE, null, objs, objs[0]);
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
						matchPartial ? 
								curDir.getName().toLowerCase().contains(folderName.toLowerCase()) : curDir.getName().equalsIgnoreCase(folderName)
								:
									matchPartial ? curDir.getName().contains(folderName) 
											:
												curDir.getName().equals(folderName)) {
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
			if (count >= Math.abs(distance)) {
				break;
			}
		}
		if (distance > 0) {
			return path.substring(index);
		} else {
			return path.substring(0,index);
		}
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
	public static void main(String args[]) {}
}